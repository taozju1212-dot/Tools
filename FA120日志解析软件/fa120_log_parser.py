"""
FA120 日志解析软件
解析全自动荧光免疫分析仪的日志，用于异常故障排查和时序检查
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re
import os
import sys
import concurrent.futures
import multiprocessing
from dataclasses import dataclass, field


def _app_dir() -> str:
    """返回配置文件所在目录：打包为 EXE 时取 EXE 旁边的目录，否则取脚本目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return _app_dir()

# ── Color palette (Material Design) ─────────────────────────────────────────
C_BG       = "#f4f7fc"   # 淡蓝白主背景
C_BG2      = "#dce8f5"   # 蓝灰工具栏/次级区域
C_CARD     = "#ffffff"   # 纯白卡片
C_BORDER   = "#b8cfe8"   # 蓝灰边框
C_BLUE     = "#1565c0"   # 深宝石蓝（主色）
C_BLUE_HV  = "#0d47a1"   # 悬停深蓝
C_GREEN    = "#2e7d32"   # 商务绿
C_RED      = "#c62828"   # 正红（报警/错误）
C_RED_LIGHT = "#ffebee"  # 淡粉（报警背景）
C_TEXT     = "#212121"   # 近黑主文字
C_TEXT2    = "#546e7a"   # 蓝灰次文字
C_WHITE    = "#ffffff"

# Timeline alternating background colors for level-1 action regions
REGION_COLORS = ["#e8f0fe", "#fef7e0", "#e6f4ea", "#fce8e6", "#f3e8fd",
                 "#e0f7fa", "#fff3e0", "#f1f8e9"]

# Mode mapping: category letter -> mode name
MODE_MAP = {
    'B': '一步法',
    'C': '单次稀释',
    'D': '多次稀释',
    'E': '两孔稀释',
}

# ── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class SampleInfo:
    serial: str          # 流水号 "0001"
    sample_id: str       # 样本ID "202603110001"
    rack_pos: str        # 样本架位置 "0-0"
    test_count: int      # 测试数
    cap_open: bool       # 是否开盖
    shake: bool          # 是否摇匀
    dilution: int        # 稀释倍数
    sample_type: str     # 样本类型
    test_items: list     # 检测项目列表
    arrange_time: str = ""
    mode: str = ""       # 测试模式 (从动作行推断)
    status: str = ""     # 样本状态：测试完成/异常/未知
    finish_time: str = ""   # 完成时间
    project_abbr: str = ""  # 项目缩写
    concentration: str = "" # 浓度
    measure_value: str = "" # 测量值
    missing_actions: list = None  # 缺失的一级动作编号列表（对比标准流程）

    def __post_init__(self):
        if self.missing_actions is None:
            self.missing_actions = []


@dataclass
class Action:
    sample_num: str      # 5位样本号 "00001"
    mode_char: str       # 模式字母 A-E
    level1: str          # 一级动作 "E00"
    level2: str          # 二级动作 "1.5"
    component: str       # 运动部件 "加样Y"
    start_pos: int       # 起始坐标
    end_pos: int         # 目标坐标
    start_time: str = "" # "HH:MM:SS.mmm"
    end_time: str = ""
    theory_raster: int = 0
    actual_raster: int = 0
    raster_deviation: int = 0
    _action_key: str = ""  # 用于关联 Start/Finish/MOTOR


@dataclass
class Alarm:
    time_str: str        # 报警时间
    error_code: str      # 错误编号 "G-INJ-012"
    sample_num: str      # 对应样本 "0005"
    action_code: str     # 对应动作 "E01"
    content: str         # 报警内容
    detail: str          # 详情


@dataclass
class InstrumentInfo:
    device_serial: str = ""
    log_time: str = ""
    user_program_version: str = ""
    request_count: int = 0
    detect_count: int = 0
    open_cap_count: int = 0
    current_arrangement_count: int = 0
    ccid: str = ""
    signal_strength: str = ""
    mid_version: str = ""
    mcu1_version: str = ""
    mcu2_version: str = ""
    mcu3_version: str = ""
    temp_control_version: str = ""


@dataclass
class SelfCheckRecord:
    check_time: str
    mcu: str
    status: str
    error_info: str = ""


@dataclass
class UserActionRecord:
    action_time: str
    action_type: str
    detail: str
    sample_serial: str = ""


# ── Log Parser ───────────────────────────────────────────────────────────────

class LogParser:
    # 样本申请
    RE_SAMPLE = re.compile(
        r'"#(\d+)-(\d{4})\s*"\s*"申请：样本架(\d+-\d+)，样本ID:(\w+),\s*测试数(\d+)，'
        r'开盖(\d)，摇匀(\d)，稀释倍数(\d+)"'
        r'.*?样本类型：\s*"([^"]+)".*?检测项目：\s*\(([^)]+)\)'
    )

    # Action Start - 捕获内嵌精确时间戳和动作信息（[M] 步进电机动作）
    RE_START = re.compile(
        r'Msg:\s*"([\d:.]+)"\s*"\[DEBUG\]:\[(\d{5})([A-Z])([A-Z]\d{2})\.(\d+\.\d+[#]*(?:\(\d+\))?)\]'
        r'action_State:\s*Start\[M\]\s*(.+?)\s+speed\[\d+\]\s+startPos\[(-?\d+)\]\s+desPos\[(-?\d+)\]"'
    )

    # Action Start - ADP 驱动动作（无部件名字段）
    RE_START_ADP = re.compile(
        r'Msg:\s*"([\d:.]+)"\s*"\[DEBUG\]:\[(\d{5})([A-Z])([A-Z]\d{2})\.(\d+\.\d+[#]*(?:\(\d+\))?)\]'
        r'action_State:\s*Start\[ADP\]\s+speed\[\d+\]\s+startPos\[(-?\d+)\]\s+desPos\[(-?\d+)\]"'
    )

    # Action Finish
    RE_FINISH = re.compile(
        r'Msg:\s*"([\d:.]+)"\s*"\[DEBUG\]:\[(\d{5})([A-Z])([A-Z]\d{2})\.(\d+\.\d+[#]*(?:\(\d+\))?)\]'
        r'action_State:Finish\[M\]\s*(.+?)\s+speed'
    )

    # Action Finish - ADP
    RE_FINISH_ADP = re.compile(
        r'Msg:\s*"([\d:.]+)"\s*"\[DEBUG\]:\[(\d{5})([A-Z])([A-Z]\d{2})\.(\d+\.\d+[#]*(?:\(\d+\))?)\]'
        r'action_State:Finish\[ADP\]\s+speed'
    )

    # MOTOR line
    RE_MOTOR = re.compile(
        r'\[MOTOR\]:\[(\d{5})([A-Z])([A-Z]\d{2})\.(\d+\.\d+)\((\d+)\)\]\s*(.+?)::'
        r'.*?理论光栅次数:(-?\d+)\s+实际光栅次数:(-?\d+)\s+光栅偏差:(-?\d+)'
    )

    # 报警  (样本号/模式/动作码均为可选)
    RE_ALARM = re.compile(
        r'(\d{2}:\d{2}:\d{2}).*?报警信息\s*"([A-Z]+-[A-Z]+-\d+)-(\d{4,5})?([A-Z])?([A-Z]\d{2})?"\s*'
        r'"([^"]*)"\s*"([^"]*)"\s*"([^"]*)"'
    )

    RE_TIME_PREFIX = re.compile(r'^(\d{2}:\d{2}:\d{2})')
    RE_SELF_CHECK = re.compile(r'Msg:\s*"([\d:.]+)"\s*"MCU([123])自检"')
    RE_SELF_CHECK_FAIL = re.compile(r'报警信息\s*"(G-OTH-20[234])')
    RE_DEVICE_SERIAL = re.compile(r'仪器序列号.*?"([^"]+)"')
    RE_CCID = re.compile(r'iccid\s*"([^"]+)"', re.IGNORECASE)
    RE_SIGNAL = re.compile(r'信号值\s+([^\r\n"]+)')
    RE_MID_VERSION = re.compile(r'MCU4Mid version.*?"([^"]+)"', re.IGNORECASE)
    RE_FW_VERSION = re.compile(r'MCU([1235])\s+SN:\s*".*?"\s*"([^"]+)"', re.IGNORECASE)
    RE_STATS = re.compile(r'申请样本次数[:：](\d+)\s+检测次数(\d+)\s+开盖次数(\d+)')
    RE_LOAD_PROJECT = re.compile(r'导入项目(\d+)-(\d+).*?项目:([^\n"]+?)\s+批次:([A-Za-z0-9_-]+)', re.S)
    RE_LOAD_CARTRIDGE = re.compile(r'"[^"]*?(\d+)\(\)')
    RE_TEST_RESULT = re.compile(
        r'"#\d+-(\d{4})(?:\s+-\d+)?"\s*测试完成\s+项目\s+"([^"]+)"\s+浓度\s+"([^"]+)"\s+测量值\s+(\S+)'
    )
    RE_COMPLETE_STAGE = re.compile(r'\[(\d{5})([A-Z])([A-Z]\d{2})\.(\d+\.\d+)')
    # 自检格式动作（[.X.Y]），用于提取部件名
    RE_SELFCHECK_COMP = re.compile(
        r'\[DEBUG\]:\[\.[\d.#()]+\]action_State:\s*Start\[M\]\s*(.+?)\s+speed'
    )

    # 理论时间表
    RE_THEORY = re.compile(r'^([A-Z]\d{2})[：:](.+?)(?:\((\d+)\))?\s*$')

    def __init__(self):
        self.samples: dict[str, SampleInfo] = {}      # key = serial "0001"
        self.actions: dict[str, list[Action]] = {}     # key = serial "0001"
        self.system_actions: list[Action] = []         # sample_num="00000" 的系统预置动作
        self.all_components: list[str] = []            # 日志中出现的所有部件名（去重有序）
        self.motor_aliases: dict[str, str] = {}        # 原始部件名 -> 显示名称
        self.alarms: list[Alarm] = []
        self.theory_times: dict[str, int] = {}         # key = "E00" -> ms
        self.theory_names: dict[str, str] = {}         # key = "E00" -> desc
        self.theory_display_names: dict[str, str] = {} # key = "E00" -> 表格显示名
        self.action_names: dict[str, str] = {}         # key = "E00" -> 用户自定义一级动作名称
        self._pending: dict[str, Action] = {}          # key for linking
        self.motor_names: list[str] = []               # 预设运动部件名称
        self.instrument_info = InstrumentInfo()
        self.self_checks: list[SelfCheckRecord] = []
        self.user_actions: list[UserActionRecord] = []
        self.standard_action_sequence: list[str] = []  # 兼容旧引用：当前日志样本0001的动作序列
        self.standard_sequences: dict[str, list[str]] = {}  # 模式名 -> 标准动作序列，按模式分类存储

    def load_standard_sequences(self, path: str):
        """从 JSON 文件加载各模式标准动作序列。"""
        if not os.path.exists(path):
            return
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                self.standard_sequences = {k: list(v) for k, v in data.items()}
        except Exception:
            pass

    def save_standard_sequences(self, path: str):
        """将各模式标准动作序列写入 JSON 文件。"""
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.standard_sequences, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_motor_names(self, path: str):
        """加载运动部件名称预设列表"""
        if not os.path.exists(path):
            return
        for enc in ('utf-8', 'gbk', 'gb2312', 'gb18030'):
            try:
                with open(path, encoding=enc) as f:
                    self.motor_names = [l.strip() for l in f if l.strip()]
                return
            except (UnicodeDecodeError, UnicodeError):
                continue

    def load_theory_time(self, path: str):
        """加载理论时间表"""
        if not os.path.exists(path):
            return
        for enc in ('utf-8', 'gbk', 'gb2312', 'gb18030'):
            try:
                with open(path, encoding=enc) as f:
                    lines = f.readlines()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            return
        for line in lines:
            m = self.RE_THEORY.match(line.strip())
            if m:
                code, desc, time_ms = m.group(1), m.group(2), m.group(3)
                self.theory_names[code] = desc.strip()
                if time_ms:
                    self.theory_times[code] = int(time_ms)

    def load_action_names(self, path: str):
        """加载一级动作名称预设（例如桌面的 FA120动作日志帧头 文件）"""
        if not os.path.exists(path):
            return

        # 支持文本文件（txt、csv 等）和 Excel 文件（xlsx/xls）
        _, ext = os.path.splitext(path)
        ext = ext.lower()

        lines = []
        if ext in ('.xlsx', '.xls'):
            try:
                import openpyxl
            except ImportError:
                # 如果没有 openpyxl，则不再继续
                return
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if not row:
                    continue
                # 拍平单元格到字符串
                row_text = ' '.join([str(c).strip() for c in row if c is not None and str(c).strip()])
                if row_text:
                    lines.append(row_text)
        else:
            for enc in ('utf-8', 'gbk', 'gb2312', 'gb18030'):
                try:
                    with open(path, encoding=enc) as f:
                        lines = f.readlines()
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            else:
                return

        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            # 支持多种格式：
            # 1) |E|0|ACT_GETTIP_ID1, //取TIP|
            # 2) E00: 取TIP
            # 3) E00：取TIP
            if '|' in line:
                parts = [p.strip() for p in line.split('|') if p.strip()]
                if len(parts) >= 3:
                    letter = parts[0]
                    num = parts[1]
                    code = f"{letter}{num.zfill(2)}"
                    rest = '|'.join(parts[2:])
                    if '//' in rest:
                        name = rest.split('//', 1)[1].strip().rstrip('|').strip()
                    else:
                        name = rest.strip().rstrip('|').strip()
                    if name:
                        self.action_names[code] = name
                continue

            m = re.match(r'^([A-Z]\d{2})\s*[:：]\s*(.+)$', line)
            if m:
                code, name = m.group(1), m.group(2).strip()
                if name:
                    self.action_names[code] = name

    def load_file(self, path: str):
        """解析日志文件"""
        self.samples.clear()
        self.actions.clear()
        self.alarms.clear()
        self._pending.clear()
        self.instrument_info = InstrumentInfo()
        self.self_checks.clear()
        self.user_actions.clear()

        for enc in ('utf-8', 'gbk', 'gb2312', 'gb18030'):
            try:
                with open(path, encoding=enc) as f:
                    lines = f.readlines()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            raise ValueError(f"无法解码文件: {path}")

        seen_alarms = set()
        for line in lines:
            self._parse_sample(line)
            self._parse_start(line)
            self._parse_finish(line)
            self._parse_motor(line)
            self._parse_alarm(line, seen_alarms)

        self._parse_self_checks(lines)
        self._parse_user_actions(lines)
        self._update_sample_status(lines)
        self._parse_instrument_info(lines)

        # 推断每个样本的测试模式（跳过前处理时序 A，取首个 B-E 模式动作）
        for serial, si in self.samples.items():
            acts = self.actions.get(serial, [])
            mode_char = next((a.mode_char for a in acts if a.mode_char in MODE_MAP), None)
            if mode_char:
                si.mode = MODE_MAP[mode_char]
            elif acts:
                si.mode = "未知模式"
            else:
                si.mode = "无动作数据"

            if not si.status:
                si.status = "异常" if any(a.sample_num == serial for a in self.alarms) else "未知"

        # 收集所有部件名（含自检格式 [.X.Y]，去重有序）
        seen_comp: set[str] = set()
        ordered: list[str] = []
        sources = (
            list(self.system_actions)
            + [a for acts in self.actions.values() for a in acts]
        )
        for act in sources:
            name = act.component
            if name and name not in seen_comp:
                seen_comp.add(name)
                ordered.append(name)
        for line in lines:
            m = self.RE_SELFCHECK_COMP.search(line)
            if m:
                name = m.group(1).strip()
                if name and name not in seen_comp:
                    seen_comp.add(name)
                    ordered.append(name)
        self.all_components = ordered
        self._check_action_completeness()

    def _extract_time_prefix(self, line: str) -> str:
        m = self.RE_TIME_PREFIX.match(line)
        return m.group(1) if m else ""

    @staticmethod
    def _extract_quoted_values(line: str) -> list[str]:
        return [m.group(1) for m in re.finditer(r'"([^"]*)"', line)]

    @staticmethod
    def _strip_log_prefix(line: str) -> str:
        return re.sub(r'^\d{2}:\d{2}:\d{2}:\s*\[[^\]]+\]\s*', '', line).strip()

    @staticmethod
    def _plus_one_rack_pos(rack_pos: str) -> str:
        try:
            x, y = rack_pos.split('-', 1)
            return f"{int(x) + 1}-{int(y) + 1}"
        except (ValueError, AttributeError):
            return rack_pos

    def _parse_instrument_info(self, lines: list[str]):
        info = InstrumentInfo()
        log_date = ""
        if len(lines) > 1:
            quoted = self._extract_quoted_values(lines[1])
            log_date = quoted[-1].strip() if quoted else self._strip_log_prefix(lines[1]).strip('"')
        if len(lines) > 2:
            quoted = self._extract_quoted_values(lines[2])
            raw_value = quoted[-1].strip() if quoted else self._strip_log_prefix(lines[2])
            info.user_program_version = raw_value.strip('"')

        start_time = self._extract_time_prefix(lines[0]) if lines else ""
        end_time = ""
        for line in reversed(lines):
            end_time = self._extract_time_prefix(line)
            if end_time:
                break
        if log_date and start_time and end_time:
            info.log_time = f"{log_date} {start_time} - {end_time}"
        elif log_date:
            info.log_time = log_date
        elif start_time and end_time:
            info.log_time = f"{start_time} - {end_time}"
        info.current_arrangement_count = len(self.samples)

        for line in lines:
            if info.request_count == 0 and info.detect_count == 0 and info.open_cap_count == 0:
                m = self.RE_STATS.search(line)
                if m:
                    info.request_count = int(m.group(1))
                    info.detect_count = int(m.group(2))
                    info.open_cap_count = int(m.group(3))

            if not info.device_serial:
                m = self.RE_DEVICE_SERIAL.search(line)
                if m:
                    info.device_serial = m.group(1).strip()

            if not info.ccid and "4G Info" in line and "iccid" in line.lower():
                m = self.RE_CCID.search(line)
                if m:
                    info.ccid = m.group(1).strip()

            if not info.signal_strength and "信号值" in line:
                m = self.RE_SIGNAL.search(line)
                if m:
                    info.signal_strength = m.group(1).strip()

            if not info.mid_version:
                m = self.RE_MID_VERSION.search(line)
                if m:
                    info.mid_version = m.group(1).strip()

            m = self.RE_FW_VERSION.search(line)
            if m:
                mcu_no, version = m.group(1), m.group(2).strip()
                if mcu_no == "1" and not info.mcu1_version:
                    info.mcu1_version = version
                elif mcu_no == "2" and not info.mcu2_version:
                    info.mcu2_version = version
                elif mcu_no == "3" and not info.mcu3_version:
                    info.mcu3_version = version
                elif mcu_no == "5" and not info.temp_control_version:
                    info.temp_control_version = version

        self.instrument_info = info

    def _parse_self_checks(self, lines: list[str]):
        current = None
        error_lines: list[str] = []
        fail_map = {
            "1": "G-OTH-202",
            "2": "G-OTH-203",
            "3": "G-OTH-204",
        }

        def finalize(status: str):
            nonlocal current, error_lines
            if not current:
                return
            self.self_checks.append(SelfCheckRecord(
                check_time=current["time"],
                mcu=current["mcu"],
                status=status,
                error_info="\n".join(error_lines).strip(),
            ))
            current = None
            error_lines = []

        for line in lines:
            start_match = self.RE_SELF_CHECK.search(line)
            if start_match:
                if current:
                    finalize("完成")
                inner_time = start_match.group(1).split(".", 1)[0]
                current = {
                    "mcu": f"MCU{start_match.group(2)}",
                    "time": inner_time,
                }
                error_lines = []
                continue

            if not current:
                continue

            fail_match = self.RE_SELF_CHECK_FAIL.search(line)
            if fail_match and fail_match.group(1) == fail_map.get(current["mcu"][-1]):
                if "自检错误" in line:
                    error_lines.append(line.strip())
                finalize("错误")

        if current:
            finalize("完成")

    def _parse_user_actions(self, lines: list[str]):
        pending_project = None

        for line in lines:
            time_str = self._extract_time_prefix(line)

            sample_match = self.RE_SAMPLE.search(line)
            if sample_match:
                serial = sample_match.group(2)
                sample = self.samples.get(serial)
                rack_display = self._plus_one_rack_pos(sample_match.group(3))
                item_text = sample.test_items[0] if sample and sample.test_items else ""
                detail = f"项目：{item_text}。样本ID:{sample_match.group(4)}"
                self.user_actions.append(UserActionRecord(
                    action_time=time_str,
                    action_type=f"样本架{rack_display}编排",
                    detail=detail,
                    sample_serial=serial,
                ))
                continue

            if "导入项目" in line:
                load_match = self.RE_LOAD_PROJECT.search(line.replace("\\n", "\n"))
                if load_match:
                    pending_project = {
                        "slot": load_match.group(1),
                        "project_no": load_match.group(2),
                        "project_name": load_match.group(3).strip(),
                        "batch": load_match.group(4).strip(),
                    }
                continue

            if "装载子弹夹" in line:
                load_match = self.RE_LOAD_CARTRIDGE.search(line)
                slot_text = load_match.group(1) if load_match else ""
                if slot_text:
                    try:
                        slot_display = str(int(slot_text) + 1)
                    except ValueError:
                        slot_display = slot_text
                else:
                    slot_display = ""

                detail = ""
                if pending_project:
                    detail = (
                        f"项目{pending_project['slot']}-{pending_project['project_no']}-"
                        f"{pending_project['project_name']} 批次:{pending_project['batch']}"
                    )

                self.user_actions.append(UserActionRecord(
                    action_time=time_str,
                    action_type=f"装载弹夹{slot_display}" if slot_display else "装载弹夹",
                    detail=detail,
                ))
                pending_project = None

    def _update_sample_status(self, lines: list[str]):
        # 第一阶段：通过 F07.3.3 Finish + 测试完成结果行配对，提取完成数据
        pending_serial = None  # 等待结果行的样本流水号

        for line in lines:
            stage_match = self.RE_COMPLETE_STAGE.search(line)
            if stage_match and "Finish" in line:
                sample_num = stage_match.group(1)
                level1 = stage_match.group(3)
                level2 = stage_match.group(4)
                if sample_num != "00000" and level1 == "F07" and level2.startswith("3.3"):
                    pending_serial = sample_num[1:]  # 后4位作为流水号

            result_match = self.RE_TEST_RESULT.search(line)
            if result_match:
                # Extract sample serial from test result line (group 1)
                # and other data (groups 2, 3, 4)
                sample_serial_from_line = result_match.group(1)
                project_abbr = result_match.group(2)
                concentration = result_match.group(3)
                measure_value = result_match.group(4)

                # Always prefer the serial embedded in the test result line;
                # fall back to pending_serial only when the line has no prefix
                serial_to_update = sample_serial_from_line or pending_serial

                if serial_to_update and serial_to_update in self.samples:
                    s = self.samples[serial_to_update]
                    s.status = "测试完成"
                    s.finish_time = self._extract_time_prefix(line)
                    s.project_abbr = project_abbr
                    s.concentration = concentration
                    s.measure_value = measure_value
                pending_serial = None

        # 第二阶段：未完成的样本，检查是否有对应报警
        for serial, sample in self.samples.items():
            if sample.status != "测试完成":
                if any(alarm.sample_num == serial for alarm in self.alarms):
                    sample.status = "异常"

    def _check_action_completeness(self):
        """以样本0001的一级动作编号顺序为标准流程，标记动作不全的其他样本。
        若该样本的模式已有持久化标准序列，则自动更新；否则以标准序列为准对比。
        """
        ref_serial = "0001"

        # 从样本0001提取本次日志的有序去重一级动作序列，并自动学习该模式的标准
        if ref_serial in self.actions and ref_serial in self.samples:
            seen: set[str] = set()
            sequence: list[str] = []
            for a in self.actions[ref_serial]:
                if a.level1 not in seen:
                    seen.add(a.level1)
                    sequence.append(a.level1)
            self.standard_action_sequence = sequence  # 兼容旧引用

            ref_mode = self.samples[ref_serial].mode  # e.g. "两孔稀释"
            if ref_mode and ref_mode not in self.standard_sequences:
                # 首次见到该模式，自动学习并写入（调用方负责持久化）
                self.standard_sequences[ref_mode] = sequence

        # 按每个样本的测试模式匹配对应的标准序列进行对比
        for serial, sample in self.samples.items():
            if serial == ref_serial:
                continue
            mode_standard = self.standard_sequences.get(sample.mode, [])
            if not mode_standard:
                continue
            sample_l1_set = {a.level1 for a in self.actions.get(serial, [])}
            missing = [code for code in mode_standard if code not in sample_l1_set]
            sample.missing_actions = missing
            if missing:
                sample.status = "异常"

    def _parse_sample(self, line: str):
        m = self.RE_SAMPLE.search(line)
        if not m:
            return
        batch, serial = m.group(1), m.group(2)
        items_str = m.group(10)
        items = [s.strip().strip('"').strip("'") for s in items_str.split(',')]
        si = SampleInfo(
            serial=serial,
            sample_id=m.group(4),
            rack_pos=m.group(3),
            test_count=int(m.group(5)),
            cap_open=(m.group(6) == '1'),
            shake=(m.group(7) == '1'),
            dilution=int(m.group(8)),
            sample_type=m.group(9),
            test_items=items,
            arrange_time=self._extract_time_prefix(line),
        )
        self.samples[serial] = si

    def _make_action_key(self, sample_num, mode_char, level1, level2_raw):
        """生成用于关联 Start/Finish/MOTOR 的唯一键"""
        # 清理level2: 去掉 # 和 (N) 后缀
        l2_clean = re.sub(r'[#()]|\(\d+\)', '', level2_raw).rstrip('.')
        return f"{sample_num}{mode_char}{level1}.{l2_clean}"

    def _parse_start(self, line: str):
        m = self.RE_START.search(line)
        adp = False
        if not m:
            m = self.RE_START_ADP.search(line)
            adp = True
        if not m:
            return
        timestamp = m.group(1)
        sample_num = m.group(2)
        mode_char = m.group(3)
        level1 = m.group(4)
        level2_raw = m.group(5)
        if adp:
            # RE_START_ADP 没有部件名组，坐标为组6和组7
            component = "ADP"
            start_pos = int(m.group(6))
            end_pos = int(m.group(7))
        else:
            component = m.group(6).strip()
            start_pos = int(m.group(7))
            end_pos = int(m.group(8))

        level2 = re.sub(r'[#]', '', level2_raw)

        action = Action(
            sample_num=sample_num,
            mode_char=mode_char,
            level1=level1,
            level2=level2,
            component=component,
            start_pos=start_pos,
            end_pos=end_pos,
            start_time=timestamp,
        )

        key = self._make_action_key(sample_num, mode_char, level1, level2_raw)
        action._action_key = key
        self._pending[key] = action

        if sample_num == "00000":
            # 系统预置动作：单独存储，供时间轴显示用
            self.system_actions.append(action)
            return

        # 取样本流水号后4位
        serial = sample_num[1:]  # "00001" -> "0001"
        if serial not in self.actions:
            self.actions[serial] = []
        self.actions[serial].append(action)

    def _parse_finish(self, line: str):
        m = self.RE_FINISH.search(line)
        if not m:
            m = self.RE_FINISH_ADP.search(line)
        if not m:
            return
        timestamp = m.group(1)
        sample_num = m.group(2)
        mode_char = m.group(3)
        level1 = m.group(4)
        level2_raw = m.group(5)

        key = self._make_action_key(sample_num, mode_char, level1, level2_raw)
        if key in self._pending:
            self._pending[key].end_time = timestamp

    def _parse_motor(self, line: str):
        m = self.RE_MOTOR.search(line)
        if not m:
            return
        sample_num = m.group(1)
        if sample_num == "00000":
            return
        mode_char = m.group(2)
        level1 = m.group(3)
        level2_base = m.group(4)
        sub_step = m.group(5)

        key = self._make_action_key(sample_num, mode_char, level1, level2_base)
        if key in self._pending:
            act = self._pending[key]
            act.theory_raster = int(m.group(7))
            act.actual_raster = int(m.group(8))
            act.raster_deviation = int(m.group(9))

    def _parse_alarm(self, line: str, seen: set):
        m = self.RE_ALARM.search(line)
        if not m:
            return
        time_str = m.group(1)
        error_code = m.group(2)
        sample_num_raw = m.group(3)
        mode_char = m.group(4)
        action_code = m.group(5)
        content = m.group(6)
        detail = m.group(7)

        # 去重：只去除同一秒内完全相同的重复条目（日志可能连续打印两次同一行）
        dedup_key = f"{time_str}-{error_code}-{sample_num_raw}-{action_code}"
        if dedup_key in seen:
            return
        seen.add(dedup_key)

        # 取后4位作为样本流水号（无样本号时留空）
        sample_serial = sample_num_raw[-4:] if sample_num_raw else ""
        action_code   = action_code or ""

        alarm = Alarm(
            time_str=time_str,
            error_code=error_code,
            sample_num=sample_serial,
            action_code=action_code,
            content=content,
            detail=detail,
        )
        self.alarms.append(alarm)


# ── Time utilities ───────────────────────────────────────────────────────────

def parse_log_job(path: str, preset_state: dict):
    parser = LogParser()
    parser.theory_times = dict(preset_state.get("theory_times", {}))
    parser.theory_names = dict(preset_state.get("theory_names", {}))
    parser.theory_display_names = dict(preset_state.get("theory_display_names", {}))
    parser.action_names = dict(preset_state.get("action_names", {}))
    parser.motor_names = list(preset_state.get("motor_names", []))
    parser.motor_aliases = dict(preset_state.get("motor_aliases", {}))
    parser.standard_sequences = {k: list(v)
                                 for k, v in preset_state.get("standard_sequences", {}).items()}
    parser.load_file(path)
    return parser


def time_to_ms(t: str) -> float:
    """将 HH:MM:SS.mmm 转为毫秒"""
    if not t:
        return 0.0
    parts = t.split(':')
    if len(parts) == 3:
        h, m, s_ms = parts
        s_parts = s_ms.split('.')
        s = float(s_parts[0])
        ms = float(s_parts[1]) if len(s_parts) > 1 else 0
        return (int(h) * 3600 + int(m) * 60 + s) * 1000 + ms
    return 0.0


def ms_to_time(ms_val: float) -> str:
    """将毫秒转为 HH:MM:SS.mmm"""
    total_s = ms_val / 1000.0
    h = int(total_s // 3600)
    rem = total_s - h * 3600
    m = int(rem // 60)
    s = rem - m * 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


# ── Timeline Canvas ──────────────────────────────────────────────────────────

class TimelineCanvas(tk.Frame):
    ROW_HEIGHT = 32
    LABEL_WIDTH = 120
    TIME_HEADER = 30
    MIN_PIXELS_PER_MS = 0.02
    MAX_PIXELS_PER_MS = 5.0

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_CARD)
        self.app = app
        self.components = []        # 运动部件名称列表 (Y轴顺序)
        self.actions = []           # 当前样本的动作列表
        self.alarms = []            # 当前样本的报警
        self.pixels_per_ms = 0.1    # 缩放比例
        self.time_offset = 0.0      # X轴起始时间(ms)
        self.time_range = 0.0       # X轴总时间范围(ms)
        self._drag_idx = None
        self._drag_y = 0

        # 左侧标签区 (固定)
        self.label_canvas = tk.Canvas(self, width=self.LABEL_WIDTH, bg=C_CARD,
                                      highlightthickness=0, bd=0)
        self.label_canvas.pack(side=tk.LEFT, fill=tk.Y)

        # 右侧绘图区 (可滚动)
        right_frame = tk.Frame(self, bg=C_CARD)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(right_frame, bg=C_CARD, highlightthickness=0, bd=0)
        self.h_scroll = tk.Scrollbar(right_frame, orient=tk.HORIZONTAL,
                                     command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.h_scroll.set)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # 绑定事件
        self.canvas.bind("<MouseWheel>", self._on_scroll_zoom)
        self.label_canvas.bind("<Button-1>", self._on_label_press)
        self.label_canvas.bind("<B1-Motion>", self._on_label_drag)
        self.label_canvas.bind("<ButtonRelease-1>", self._on_label_release)
        self.canvas.bind("<Configure>", lambda e: self._draw())
        self.canvas.bind("<Button-1>", self._on_canvas_click)

    def set_data(self, actions: list[Action], alarms: list[Alarm],
                 motor_names: list[str], system_actions: list[Action] = None):
        # 将样本前的系统动作（00000）追加到动作列表头部
        if system_actions and actions:
            first_ms = min(
                (time_to_ms(a.start_time) for a in actions if a.start_time),
                default=None,
            )
            if first_ms is not None:
                pre = [a for a in system_actions
                       if a.start_time and first_ms - time_to_ms(a.start_time) <= 30_000]
                actions = pre + list(actions)

        self.actions = actions
        self.alarms = alarms

        # 确定运动部件列表 (按预设顺序，补充日志中新出现的)
        seen = set()
        self.components = []
        for name in motor_names:
            if name not in seen:
                self.components.append(name)
                seen.add(name)
        for act in actions:
            if act.component not in seen:
                self.components.append(act.component)
                seen.add(act.component)

        # 计算时间范围
        times = []
        for act in actions:
            if act.start_time:
                times.append(time_to_ms(act.start_time))
            if act.end_time:
                times.append(time_to_ms(act.end_time))
        if times:
            self.time_offset = min(times) - 500
            self.time_range = max(times) - self.time_offset + 500
        else:
            self.time_offset = 0
            self.time_range = 1000

        # 自适应缩放: 尝试让内容填满可见区域
        canvas_w = max(self.canvas.winfo_width(), 600)
        self.pixels_per_ms = canvas_w / self.time_range if self.time_range > 0 else 0.1
        self.pixels_per_ms = max(self.MIN_PIXELS_PER_MS,
                                 min(self.MAX_PIXELS_PER_MS, self.pixels_per_ms))
        self._draw()

    def _draw(self):
        self.canvas.delete("all")
        self.label_canvas.delete("all")

        if not self.components:
            return

        total_w = max(int(self.time_range * self.pixels_per_ms), 800)
        total_h = self.TIME_HEADER + len(self.components) * self.ROW_HEIGHT + 10
        self.canvas.configure(scrollregion=(0, 0, total_w, total_h))
        self.label_canvas.configure(height=total_h)

        # ── 背景色区域 (按一级动作) ──
        self._draw_regions(total_h)

        # ── 时间刻度 ──
        self._draw_time_axis(total_w)

        # ── Y轴标签 ──
        aliases = self.app.parser.motor_aliases
        for i, comp in enumerate(self.components):
            y = self.TIME_HEADER + i * self.ROW_HEIGHT + self.ROW_HEIGHT // 2
            label = aliases.get(comp, comp)
            self.label_canvas.create_text(
                self.LABEL_WIDTH - 5, y, text=label, anchor="e",
                font=("Microsoft YaHei", 8), fill=C_TEXT)
            # 水平基线
            base_y = self.TIME_HEADER + i * self.ROW_HEIGHT + self.ROW_HEIGHT - 4
            self.canvas.create_line(0, base_y, total_w, base_y,
                                    fill=C_BORDER, dash=(2, 4), tags="grid")

        # ── 动作矩形 (高电平信号) ──
        comp_colors = {}
        color_list = ["#4285f4", "#34a853", "#fbbc04", "#ea4335", "#8e24aa",
                      "#00acc1", "#ff7043", "#5c6bc0", "#26a69a", "#d81b60"]
        for act in self.actions:
            if not act.start_time or not act.end_time:
                continue
            if act.component not in comp_colors:
                idx = len(comp_colors) % len(color_list)
                comp_colors[act.component] = color_list[idx]
            if act.component not in self.components:
                continue

            row = self.components.index(act.component)
            x1 = (time_to_ms(act.start_time) - self.time_offset) * self.pixels_per_ms
            x2 = (time_to_ms(act.end_time) - self.time_offset) * self.pixels_per_ms
            if x2 - x1 < 1:
                x2 = x1 + 1

            y_top = self.TIME_HEADER + row * self.ROW_HEIGHT + 4
            y_bot = self.TIME_HEADER + row * self.ROW_HEIGHT + self.ROW_HEIGHT - 4
            y_base = y_bot

            color = comp_colors[act.component]
            # 高电平矩形
            self.canvas.create_rectangle(x1, y_top, x2, y_base,
                                         fill=color, outline=color,
                                         tags=("action",))
            # 左右竖线 (上升/下降沿)
            self.canvas.create_line(x1, y_base, x1, y_top, fill=color, width=1)
            self.canvas.create_line(x2, y_top, x2, y_base, fill=color, width=1)

        # ── 报警红色标记 ──
        for alarm in self.alarms:
            alarm_ms = time_to_ms(alarm.time_str + ".000") if '.' not in alarm.time_str else time_to_ms(alarm.time_str)
            x = (alarm_ms - self.time_offset) * self.pixels_per_ms
            self.canvas.create_line(x, self.TIME_HEADER, x, total_h,
                                    fill=C_RED, width=2, tags=("alarm",))
            self.canvas.create_text(x, self.TIME_HEADER - 2,
                                    text=alarm.error_code, anchor="s",
                                    font=("Microsoft YaHei", 7), fill=C_RED,
                                    tags=("alarm",))

    def _draw_regions(self, total_h):
        """按一级动作绘制交替背景色区域"""
        if not self.actions:
            return
        # 按一级动作分组，找到时间区间
        level1_ranges = {}
        for act in self.actions:
            if not act.start_time:
                continue
            t_start = time_to_ms(act.start_time)
            t_end = time_to_ms(act.end_time) if act.end_time else t_start
            if act.level1 not in level1_ranges:
                level1_ranges[act.level1] = [t_start, t_end]
            else:
                level1_ranges[act.level1][0] = min(level1_ranges[act.level1][0], t_start)
                level1_ranges[act.level1][1] = max(level1_ranges[act.level1][1], t_end)

        sorted_levels = sorted(level1_ranges.items(), key=lambda x: x[1][0])
        for i, (level1, (t_min, t_max)) in enumerate(sorted_levels):
            x1 = (t_min - self.time_offset) * self.pixels_per_ms
            x2 = (t_max - self.time_offset) * self.pixels_per_ms
            color = REGION_COLORS[i % len(REGION_COLORS)]
            self.canvas.create_rectangle(x1, 0, x2, total_h,
                                         fill=color, outline="", tags="region")
            # 在顶部标注一级动作编号
            mid_x = (x1 + x2) / 2
            self.canvas.create_text(mid_x, 8, text=level1,
                                    font=("Microsoft YaHei", 7, "bold"),
                                    fill=C_TEXT2, tags="region_label")
        # 把 region 放到最底层
        self.canvas.tag_lower("region")

    def _draw_time_axis(self, total_w):
        """绘制时间刻度"""
        # 根据缩放级别选择合适的刻度间隔
        target_px = 80  # 目标每个刻度间隔像素数
        interval_ms = target_px / self.pixels_per_ms
        # 取整到合适的数值
        nice_intervals = [100, 200, 500, 1000, 2000, 5000, 10000, 30000, 60000]
        for ni in nice_intervals:
            if ni >= interval_ms:
                interval_ms = ni
                break
        else:
            interval_ms = nice_intervals[-1]

        start_ms = (int(self.time_offset / interval_ms)) * interval_ms
        t = start_ms
        while t < self.time_offset + self.time_range:
            x = (t - self.time_offset) * self.pixels_per_ms
            if x >= 0:
                self.canvas.create_line(x, self.TIME_HEADER - 5, x, self.TIME_HEADER,
                                        fill=C_TEXT2, tags="axis")
                label = ms_to_time(t)
                # 显示 HH:MM:SS
                short_label = label[:8]
                self.canvas.create_text(x, self.TIME_HEADER - 7, text=short_label,
                                        anchor="s", font=("Consolas", 7),
                                        fill=C_TEXT2, tags="axis")
            t += interval_ms

    def _on_scroll_zoom(self, event):
        """鼠标滚轮缩放X轴"""
        factor = 1.2 if event.delta > 0 else 1 / 1.2
        new_scale = self.pixels_per_ms * factor
        new_scale = max(self.MIN_PIXELS_PER_MS, min(self.MAX_PIXELS_PER_MS, new_scale))
        self.pixels_per_ms = new_scale
        self._draw()

    def _on_canvas_click(self, event):
        """点击画布，检查是否点中报警线"""
        cx = self.canvas.canvasx(event.x)
        items = self.canvas.find_closest(cx, event.y)
        if items:
            tags = self.canvas.gettags(items[0])
            if "alarm" in tags:
                # 找到最近的报警
                for alarm in self.alarms:
                    t = time_to_ms(alarm.time_str + ".000") if '.' not in alarm.time_str else time_to_ms(alarm.time_str)
                    ax = (t - self.time_offset) * self.pixels_per_ms
                    if abs(ax - cx) < 10:
                        self.app.highlight_alarm(alarm)
                        break

    # ── Y轴拖拽排序 ──
    def _on_label_press(self, event):
        y = event.y - self.TIME_HEADER
        if y < 0:
            return
        idx = y // self.ROW_HEIGHT
        if 0 <= idx < len(self.components):
            self._drag_idx = idx
            self._drag_y = event.y

    def _on_label_drag(self, event):
        if self._drag_idx is None:
            return
        dy = event.y - self._drag_y
        if abs(dy) >= self.ROW_HEIGHT:
            new_idx = self._drag_idx + (1 if dy > 0 else -1)
            if 0 <= new_idx < len(self.components):
                self.components[self._drag_idx], self.components[new_idx] = \
                    self.components[new_idx], self.components[self._drag_idx]
                self._drag_idx = new_idx
                self._drag_y = event.y
                self._draw()

    def _on_label_release(self, event):
        self._drag_idx = None

    def scroll_to_time(self, time_ms: float):
        """滚动到指定时间位置"""
        if self.time_range <= 0:
            return
        frac = (time_ms - self.time_offset) / self.time_range
        frac = max(0, min(1, frac))
        self.canvas.xview_moveto(max(0, frac - 0.1))


# ── Table View ───────────────────────────────────────────────────────────────

class TableView(tk.Frame):
    FIXED_COLS = ["动作编号", "动作名称", "二级动作", "运动部件"]
    OPTIONAL_COLS = ["动作坐标", "起始时间", "结束时间", "动作时间(ms)",
                     "理论光栅", "实际光栅", "光栅偏差"]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_CARD)
        self.app = app
        self.col_vars = {}
        self._collapse_var = tk.BooleanVar(value=False)

        # 列显示/隐藏 控制栏
        ctrl_frame = tk.Frame(self, bg=C_BG2)
        ctrl_frame.pack(side=tk.TOP, fill=tk.X, padx=2, pady=2)
        tk.Label(ctrl_frame, text="显示列:", font=("Microsoft YaHei", 8),
                 bg=C_BG2, fg=C_TEXT2).pack(side=tk.LEFT, padx=4)

        for col in self.OPTIONAL_COLS:
            var = tk.BooleanVar(value=True)
            self.col_vars[col] = var
            cb = ttk.Checkbutton(ctrl_frame, text=col, variable=var,
                                 command=self._rebuild_columns)
            cb.pack(side=tk.LEFT, padx=2)

        # 分隔线
        ttk.Separator(ctrl_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        ttk.Checkbutton(ctrl_frame, text="仅显示一级动作", variable=self._collapse_var,
                        command=self._toggle_collapse).pack(side=tk.LEFT, padx=2)

        # Treeview
        tree_frame = tk.Frame(self, bg=C_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # 禁止用户拖拽调整列宽
        self.tree.bind("<Button-1>", self._block_col_resize)

        self._rebuild_columns()

    def _block_col_resize(self, event):
        if self.tree.identify_region(event.x, event.y) == "separator":
            return "break"

    def _auto_resize_columns(self):
        import tkinter.font as tkfont
        font = tkfont.Font(font=("Microsoft YaHei", 9))
        cols = self._get_visible_cols()
        widths = {col: font.measure(col) + 24 for col in cols}
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            for col, val in zip(cols, vals):
                w = font.measure(str(val)) + 24
                if w > widths[col]:
                    widths[col] = w
        for col in cols:
            self.tree.column(col, width=max(60, widths[col]))

    def _get_visible_cols(self):
        cols = list(self.FIXED_COLS)
        for col in self.OPTIONAL_COLS:
            if self.col_vars.get(col, tk.BooleanVar(value=True)).get():
                cols.append(col)
        return cols

    def _format_level1(self, level1: str) -> str:
        """根据理论时间设置中的动作描述格式化一级动作显示文本"""
        name = self.app.parser.theory_names.get(level1, "")
        return f"{name}:{level1}" if name else level1

    def _rebuild_columns(self):
        cols = self._get_visible_cols()
        self.tree["columns"] = cols
        col_widths = {
            "动作编号": 60, "动作名称": 130, "二级动作": 220, "运动部件": 80,
            "动作坐标": 110, "起始时间": 85, "结束时间": 85,
            "动作时间(ms)": 90, "理论光栅": 70, "实际光栅": 70, "光栅偏差": 70,
        }
        for col in cols:
            w = col_widths.get(col, 80)
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col, width=w, minwidth=50, anchor="center")
        self._repopulate()

    def _toggle_collapse(self):
        """展开或折叠所有父行子节点"""
        expanded = not self._collapse_var.get()
        for parent in self.tree.get_children():
            self.tree.item(parent, open=expanded)

    def set_data(self, actions: list[Action], alarms: list[Alarm],
                 theory_times: dict, standard_sequence: list = None,
                 missing_actions: list = None):
        self._actions = actions
        self._alarms = alarms
        self._theory_times = theory_times
        self._standard_sequence = standard_sequence or []
        self._missing_actions = set(missing_actions) if missing_actions else set()
        self._repopulate()

    def _repopulate(self):
        self.tree.delete(*self.tree.get_children())
        if not hasattr(self, '_actions'):
            return

        cols = self._get_visible_cols()
        alarm_actions = {a.action_code for a in self._alarms}

        # ── 按时间顺序构建执行批次 ──
        batches = self._build_ordered_batches(self._actions)
        batch_map: dict[str, list] = {lv1: grp for lv1, grp in batches}
        # 已插入的level1集合，用于识别标准流程之外的额外动作
        inserted_l1: set[str] = set()

        def _insert_batch(level1, group):
            """插入一个动作批次的父行和所有子行。"""
            starts = [time_to_ms(a.start_time) for a in group if a.start_time]
            ends   = [time_to_ms(a.end_time)   for a in group if a.end_time]
            total_dur = round(max(ends) - min(starts), 1) if (starts and ends) else ""
            theory  = self._theory_times.get(level1, "")
            timeout = None
            if isinstance(theory, int) and total_dur != "":
                timeout = round(total_dur - theory, 1)

            theory_text = f"理论：{theory}ms" if isinstance(theory, int) else "理论：--"
            if timeout is None:
                timeout_text = "超时：--"
            else:
                sign = "+" if timeout > 0 else ""
                timeout_text = f"超时：{sign}{timeout}ms"

            level1_name = (self.app.parser.theory_display_names.get(level1)
                           or self.app.parser.theory_names.get(level1, ""))
            parent_row = []
            for col in cols:
                if col == "动作编号":
                    parent_row.append(level1)
                elif col == "动作名称":
                    parent_row.append(level1_name)
                elif col == "二级动作":
                    parent_row.append(theory_text)
                elif col == "运动部件":
                    parent_row.append(timeout_text)
                elif col == "动作坐标":
                    parent_row.append("")
                elif col == "起始时间":
                    parent_row.append(ms_to_time(min(starts)) if starts else "")
                elif col == "结束时间":
                    parent_row.append(ms_to_time(max(ends)) if ends else "")
                elif col == "动作时间(ms)":
                    parent_row.append(total_dur)
                else:
                    parent_row.append("")

            # 超时：仅改变文字颜色（橙色），不整行标红；缺失行才整行标红
            parent_tag = "timeout_summary" if (timeout is not None and timeout > 0) else "summary"
            pid = self.tree.insert("", tk.END, values=parent_row,
                                   tags=(parent_tag,), open=not self._collapse_var.get())

            for act in group:
                t_start = time_to_ms(act.start_time) if act.start_time else 0
                t_end   = time_to_ms(act.end_time)   if act.end_time   else 0
                duration = round(t_end - t_start, 1) if (t_start and t_end) else ""

                child_row = []
                for col in cols:
                    if col in ("动作编号", "动作名称"):
                        child_row.append("")
                    elif col == "二级动作":
                        child_row.append(act.level2)
                    elif col == "运动部件":
                        aliases = self.app.parser.motor_aliases
                        child_row.append(aliases.get(act.component, act.component))
                    elif col == "动作坐标":
                        child_row.append(f"{act.start_pos}→{act.end_pos}")
                    elif col == "起始时间":
                        child_row.append(act.start_time)
                    elif col == "结束时间":
                        child_row.append(act.end_time)
                    elif col == "动作时间(ms)":
                        child_row.append(duration)
                    elif col == "理论光栅":
                        child_row.append(act.theory_raster)
                    elif col == "实际光栅":
                        child_row.append(act.actual_raster)
                    elif col == "光栅偏差":
                        child_row.append(act.raster_deviation)
                    else:
                        child_row.append("")

                ciid = self.tree.insert(pid, tk.END, values=child_row)
                tags = []
                if act.level1 in alarm_actions:
                    tags.append("alarm")
                if abs(act.raster_deviation) >= 3:
                    tags.append("high_dev")
                if tags:
                    self.tree.item(ciid, tags=tuple(tags))

        def _insert_missing_row(level1):
            """插入一个占位红色行，表示该动作在标准流程中存在但本样本缺失。"""
            level1_name = (self.app.parser.theory_display_names.get(level1)
                           or self.app.parser.theory_names.get(level1, ""))
            missing_row = []
            for col in cols:
                if col == "动作编号":
                    missing_row.append(level1)
                elif col == "动作名称":
                    missing_row.append(level1_name)
                elif col == "二级动作":
                    missing_row.append("【缺失】")
                else:
                    missing_row.append("")
            self.tree.insert("", tk.END, values=missing_row, tags=("missing_action",))

        # ── 按标准流程顺序插入：有则显示，缺则插红色占位行 ──
        if self._standard_sequence:
            for level1 in self._standard_sequence:
                if level1 in batch_map:
                    _insert_batch(level1, batch_map[level1])
                elif level1 in self._missing_actions:
                    _insert_missing_row(level1)
                inserted_l1.add(level1)
            # 插入不在标准流程中的额外动作批次（按原始顺序）
            for level1, group in batches:
                if level1 not in inserted_l1:
                    _insert_batch(level1, group)
        else:
            # 无标准流程时按原始时间顺序插入
            for level1, group in batches:
                _insert_batch(level1, group)

        _bold = ("Microsoft YaHei", 9, "bold")
        self.tree.tag_configure("alarm", background=C_RED_LIGHT)
        self.tree.tag_configure("high_dev", foreground=C_RED)
        self.tree.tag_configure("summary",
                                background="#c8dcfa", font=_bold)
        # 超时：仅改橙色文字，行背景保持蓝色，与缺失行的整行红色区分
        self.tree.tag_configure("timeout_summary",
                                background="#c8dcfa", foreground="#cc6600", font=_bold)
        # 缺失动作：整行红色背景，醒目提示
        self.tree.tag_configure("missing_action",
                                background="#ffd6d6", foreground=C_RED, font=_bold)
        self._auto_resize_columns()

    @staticmethod
    def _build_ordered_batches(actions: list[Action]) -> list[tuple[str, list[Action]]]:
        """按时间顺序构建执行批次。
        同一level1的连续子步骤归为一批，穿插的其他level1提取到当前批次之后。
        例：E00, E00, D00, E00, E00 → (E00, [4条]), (D00, [1条])
        """
        if not actions:
            return []

        result = []
        i = 0
        n = len(actions)

        while i < n:
            dominant = actions[i].level1
            batch = [actions[i]]
            interleaved = []  # 被穿插的其他level1动作
            i += 1

            while i < n:
                act = actions[i]
                if act.level1 == dominant:
                    # 同一level1 → 加入当前批次
                    batch.append(act)
                    i += 1
                else:
                    # 不同level1 → 向前查找当前dominant是否还会继续出现
                    found_more = False
                    for k in range(i + 1, min(i + 8, n)):
                        if actions[k].level1 == dominant:
                            found_more = True
                            break
                    if found_more:
                        # dominant还会出现 → 这个action是穿插的，暂存
                        interleaved.append(act)
                        i += 1
                    else:
                        # dominant不再出现 → 当前批次结束
                        break

            # 输出当前dominant批次
            result.append((dominant, batch))
            # 输出穿插的动作，按各自level1分组
            if interleaved:
                il_groups: dict[str, list[Action]] = {}
                il_order: list[str] = []
                for a in interleaved:
                    if a.level1 not in il_groups:
                        il_groups[a.level1] = []
                        il_order.append(a.level1)
                    il_groups[a.level1].append(a)
                for il1 in il_order:
                    result.append((il1, il_groups[il1]))

        return result

    def highlight_action(self, action_code: str):
        """高亮指定动作编号对应的父行（展开并滚动到视图）"""
        for parent in self.tree.get_children():
            vals = self.tree.item(parent, "values")
            if not vals:
                continue
            # vals[0] 即「动作编号」列
            if vals[0] == action_code:
                self.tree.item(parent, open=True)
                self.tree.selection_set(parent)
                self.tree.see(parent)
                return


# ── Theory Time Dialog ───────────────────────────────────────────────────────

class TheoryTimeDialog(tk.Toplevel):
    def __init__(self, parent, parser: LogParser, on_change=None, save_path: str = ""):
        super().__init__(parent)
        self.parser = parser
        self.on_change = on_change
        self.save_path = save_path
        self.title("动作设置")
        self.geometry("680x500")
        self.configure(bg=C_BG)

        # 按钮栏
        btn_frame = tk.Frame(self, bg=C_BG)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(btn_frame, text="载入文件", command=self._load_file,
                  bg=C_BLUE, fg=C_WHITE, relief="flat", padx=10).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="保存文件", command=self._save_file,
                  bg=C_GREEN, fg=C_WHITE, relief="flat", padx=10).pack(side=tk.LEFT, padx=5)

        # 编辑表格
        tree_frame = tk.Frame(self, bg=C_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("动作编号", "动作描述", "表格显示名", "理论时间(ms)")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                 selectmode="browse")
        col_widths = {"动作编号": 70, "动作描述": 220, "表格显示名": 160, "理论时间(ms)": 110}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=col_widths.get(c, 120))
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", self._on_double_click)

        self._populate()

    def _auto_save(self):
        if not self.save_path:
            return
        import json
        data = {
            "theory_names": self.parser.theory_names,
            "theory_display_names": self.parser.theory_display_names,
            "theory_times": self.parser.theory_times,
        }
        with open(self.save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        all_codes = sorted(set(list(self.parser.theory_names.keys()) +
                              list(self.parser.theory_times.keys())))
        for code in all_codes:
            desc     = self.parser.theory_names.get(code, "")
            display  = self.parser.theory_display_names.get(code, "")
            time_ms  = self.parser.theory_times.get(code, "")
            self.tree.insert("", tk.END, values=(code, desc, display, time_ms))

    def _on_double_click(self, event):
        item = self.tree.focus()
        if not item:
            return
        col = self.tree.identify_column(event.x)
        if col not in ("#2", "#3", "#4"):
            return
        vals = self.tree.item(item, "values")
        code = vals[0]

        if col == "#2":  # 动作描述
            new_val = simpledialog.askstring(
                "编辑动作描述", f"请输入 {code} 的动作描述:",
                initialvalue=str(vals[1]), parent=self)
            if new_val is not None:
                if new_val.strip():
                    self.parser.theory_names[code] = new_val.strip()
                else:
                    self.parser.theory_names.pop(code, None)
                self._auto_save()
                self._populate()
                if self.on_change:
                    self.on_change()
        elif col == "#3":  # 表格显示名
            new_val = simpledialog.askstring(
                "编辑表格显示名", f"请输入 {code} 在表格中的显示名称（留空则用动作描述）:",
                initialvalue=str(vals[2]), parent=self)
            if new_val is not None:
                if new_val.strip():
                    self.parser.theory_display_names[code] = new_val.strip()
                else:
                    self.parser.theory_display_names.pop(code, None)
                self._auto_save()
                self._populate()
                if self.on_change:
                    self.on_change()
        else:  # col == "#4"，理论时间
            new_val = simpledialog.askstring(
                "编辑理论时间", f"请输入 {code} 的理论时间(ms):",
                initialvalue=str(vals[3]), parent=self)
            if new_val is not None:
                try:
                    self.parser.theory_times[code] = int(new_val)
                except ValueError:
                    if new_val == "":
                        self.parser.theory_times.pop(code, None)
                self._auto_save()
                self._populate()
                if self.on_change:
                    self.on_change()

    def _load_file(self):
        path = filedialog.askopenfilename(
            title="载入理论时间表",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
        if path:
            self.parser.load_theory_time(path)
            self._auto_save()
            self._populate()
            if self.on_change:
                self.on_change()

    def _save_file(self):
        path = filedialog.asksaveasfilename(
            title="保存理论时间表", defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")])
        if path:
            all_codes = sorted(set(list(self.parser.theory_names.keys()) +
                                  list(self.parser.theory_times.keys())))
            with open(path, 'w', encoding='utf-8') as f:
                for code in all_codes:
                    desc = self.parser.theory_names.get(code, "")
                    time_ms = self.parser.theory_times.get(code, None)
                    if time_ms is not None:
                        f.write(f"{code}：{desc}({time_ms})\n")
                    else:
                        f.write(f"{code}：{desc}\n")
            if self.on_change:
                self.on_change()


class MotorNamesDialog(tk.Toplevel):
    """电机名称设置：为每个部件设置 Y 轴显示名称"""

    # 固定包含的 6 个部件（即使日志中未出现也始终显示）
    REQUIRED_COMPONENTS = ["进卡Y", "退卡Y", "样本检测Y", "进样本架Y", "退样本架Y", "摆渡车X"]

    def __init__(self, parent, parser: LogParser, on_change=None, save_path: str = ""):
        super().__init__(parent)
        self.parser = parser
        self.on_change = on_change
        self.save_path = save_path
        self.title("电机名称设置")
        self.geometry("480x520")
        self.configure(bg=C_BG)

        tk.Label(self, text="双击「显示名称」列可修改，留空则使用原始名称",
                 bg=C_BG, fg=C_TEXT2, font=("Microsoft YaHei", 8)
                 ).pack(anchor="w", padx=12, pady=(8, 2))

        # 编辑表格
        tree_frame = tk.Frame(self, bg=C_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("原始名称", "显示名称")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                 selectmode="browse")
        self.tree.heading("原始名称", text="原始名称（日志中）")
        self.tree.heading("显示名称", text="显示名称（Y轴）")
        self.tree.column("原始名称", width=200)
        self.tree.column("显示名称", width=200)
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", self._on_double_click)

        # 底部按钮
        btn_frame = tk.Frame(self, bg=C_BG)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)
        tk.Button(btn_frame, text="清除所有别名", command=self._clear_all,
                  bg=C_RED, fg=C_WHITE, relief="flat", padx=10).pack(side=tk.LEFT)

        self._populate()

    def _all_components(self) -> list[str]:
        """合并日志部件 + 必须包含的 6 个部件，去重保序"""
        seen: set[str] = set()
        result: list[str] = []
        for name in self.parser.all_components + self.REQUIRED_COMPONENTS:
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        for raw in self._all_components():
            alias = self.parser.motor_aliases.get(raw, "")
            self.tree.insert("", tk.END, iid=raw, values=(raw, alias))

    def _save(self):
        if not self.save_path:
            return
        import json
        with open(self.save_path, "w", encoding="utf-8") as f:
            json.dump(self.parser.motor_aliases, f, ensure_ascii=False, indent=2)

    def _on_double_click(self, event):
        item = self.tree.focus()
        if not item:
            return
        col = self.tree.identify_column(event.x)
        if col != "#2":
            return
        raw = self.tree.item(item, "values")[0]
        current = self.parser.motor_aliases.get(raw, "")
        new_val = simpledialog.askstring(
            "编辑显示名称", f"请输入「{raw}」的显示名称（留空恢复原始名称）:",
            initialvalue=current, parent=self)
        if new_val is not None:
            if new_val.strip():
                self.parser.motor_aliases[raw] = new_val.strip()
            else:
                self.parser.motor_aliases.pop(raw, None)
            self.tree.item(item, values=(raw, self.parser.motor_aliases.get(raw, "")))
            self._save()
            if self.on_change:
                self.on_change()

    def _clear_all(self):
        self.parser.motor_aliases.clear()
        self._save()
        self._populate()
        if self.on_change:
            self.on_change()


# ── Params View ──────────────────────────────────────────────────────────────

class ParamsView(tk.Frame):
    """整机参数导入与对比视图（新增 Tab）"""

    SYSTEM_KEYS = [
        "SN_MCU0","SN_MCU1","SN_MCU2","SN_MCU3","SN_MCU4","SN_MCU5",
        "Version_MCU0","Version_MCU1","Version_MCU2","Version_MCU3",
        "Version_MCU4","Version_MCU5","Version_Needle","Version_test",
    ]

    def __init__(self, parent, app_dir: str):
        super().__init__(parent, bg=C_CARD)
        self.app_dir = app_dir
        self.data = [None, None]          # 两份参数 dict
        self.file_names = ["", ""]
        self.diff_only = tk.BooleanVar(value=False)
        self._diff_count = 0
        self._leaf_iids: list[tuple[str, bool]] = []   # (iid, is_diff)
        self._detached_items: list[tuple[str, str, int]] = []  # (iid, parent, index)
        self._action_names: dict[str, str] = {}        # "motor_i" -> "名称"
        self._param_notes: dict[str, str] = {}         # label -> 备注说明
        self._search_var = tk.StringVar()
        self._build_ui()
        self._load_action_names()
        self._load_param_notes()

    # ── UI 构建 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 顶部工具栏
        top = tk.Frame(self, bg=C_BG2)
        top.pack(fill=tk.X, padx=6, pady=4)

        # 参数1
        tk.Button(top, text="载入参数1", command=lambda: self._load_file(0),
                  bg=C_BLUE, fg=C_WHITE, relief="flat", padx=10
                  ).pack(side=tk.LEFT, padx=(0, 4))
        self._lbl1 = tk.Label(top, text="（未载入）", bg=C_BG2, fg=C_TEXT2,
                              font=("Microsoft YaHei", 8))
        self._lbl1.pack(side=tk.LEFT, padx=(0, 14))

        # 参数2
        tk.Button(top, text="载入参数2", command=lambda: self._load_file(1),
                  bg=C_BLUE, fg=C_WHITE, relief="flat", padx=10
                  ).pack(side=tk.LEFT, padx=(0, 4))
        self._lbl2 = tk.Label(top, text="（未载入）", bg=C_BG2, fg=C_TEXT2,
                              font=("Microsoft YaHei", 8))
        self._lbl2.pack(side=tk.LEFT, padx=(0, 14))

        # 仅显示差异 + 差异计数
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        ttk.Checkbutton(top, text="仅显示差异", variable=self.diff_only,
                        command=self._apply_filter).pack(side=tk.LEFT)
        self._diff_lbl = tk.Label(top, text="差异数：—", bg=C_BG2, fg=C_TEXT2,
                                  font=("Microsoft YaHei", 8))
        self._diff_lbl.pack(side=tk.LEFT, padx=8)

        # 搜索栏
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(top, text="搜索:", bg=C_BG2, fg=C_TEXT2,
                 font=("Microsoft YaHei", 8)).pack(side=tk.LEFT)
        search_entry = ttk.Entry(top, textvariable=self._search_var, width=18)
        search_entry.pack(side=tk.LEFT, padx=(4, 2))
        search_entry.bind("<Return>", lambda e: self._do_search())
        self._search_var.trace("w", lambda *a: self._do_search())
        tk.Button(top, text="✕", command=self._clear_search,
                  bg=C_BG2, fg=C_TEXT2, relief="flat", font=("Microsoft YaHei", 8),
                  cursor="hand2", padx=2).pack(side=tk.LEFT)

        # Treeview
        tree_frame = tk.Frame(self, bg=C_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        cols = ("参数文件1", "参数文件2", "备注")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="tree headings",
                                 selectmode="browse")
        self.tree.heading("#0", text="参数名称", anchor="w")
        self.tree.column("#0", width=260, minwidth=140, stretch=True)
        self.tree.heading("参数文件1", text="参数文件1", anchor="w")
        self.tree.heading("参数文件2", text="参数文件2", anchor="w")
        self.tree.heading("备注", text="备注说明（双击编辑）", anchor="w")
        self.tree.column("参数文件1", width=180, minwidth=80, anchor="w")
        self.tree.column("参数文件2", width=180, minwidth=80, anchor="w")
        self.tree.column("备注", width=200, minwidth=80, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # tag 样式
        self.tree.tag_configure("diff",   background="#ffd6d6")
        self.tree.tag_configure("missing", background="#fff3cd")
        self.tree.tag_configure("group",   background="#eef2fa",
                                font=("Microsoft YaHei", 9, "bold"))
        self.tree.tag_configure("group_diff", background="#ffe8e8",
                                font=("Microsoft YaHei", 9, "bold"))

        self.tree.bind("<Double-1>", self._on_double_click)

    # ── 文件加载 ─────────────────────────────────────────────────────────────

    def _load_file(self, idx: int):
        path = filedialog.askopenfilename(
            title=f"载入参数文件{idx+1}",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        import json as _json
        with open(path, encoding="utf-8") as f:
            self.data[idx] = _json.load(f)
        name = os.path.basename(path)
        self.file_names[idx] = name
        [self._lbl1, self._lbl2][idx].config(text=name)
        self._populate()

    # ── 树填充 ───────────────────────────────────────────────────────────────

    def _v(self, d, *keys):
        """安全取嵌套值，不存在返回 None"""
        obj = d
        for k in keys:
            if not isinstance(obj, dict):
                return None
            obj = obj.get(k)
        return obj

    def _fmt(self, v) -> str:
        if v is None:
            return "—"
        if isinstance(v, float):
            return str(round(v, 6))
        if isinstance(v, list):
            return str(v)
        return str(v)

    def _is_diff(self, v1, v2) -> bool:
        if v1 is None or v2 is None:
            return True
        if isinstance(v1, float) or isinstance(v2, float):
            return round(float(v1), 6) != round(float(v2), 6)
        return str(v1) != str(v2)

    def _insert_group(self, parent, label: str, open_: bool = True) -> str:
        return self.tree.insert(parent, tk.END, text=label, values=("", ""),
                                tags=("group",), open=open_)

    def _insert_leaf(self, parent: str, label: str, v1, v2) -> bool:
        """插入一行叶节点，返回 is_diff"""
        s1, s2 = self._fmt(v1), self._fmt(v2)
        if v1 is None or v2 is None:
            tag = "missing"
            diff = True
        else:
            diff = self._is_diff(v1, v2)
            tag = "diff" if diff else ""
        note = self._param_notes.get(label, "")
        iid = self.tree.insert(parent, tk.END, text=label, values=(s1, s2, note),
                               tags=(tag,) if tag else ())
        self._leaf_iids.append((iid, diff))
        return diff

    def _mark_group_diff(self, gid: str, has_diff: bool):
        if has_diff:
            self.tree.item(gid, tags=("group_diff",))

    def _populate(self):
        self._reattach_all()                        # 先还原所有 detach 的节点
        self.tree.delete(*self.tree.get_children())
        self._leaf_iids = []
        self._diff_count = 0

        d1 = self.data[0] or {}
        d2 = self.data[1] or {}
        all_keys = list(dict.fromkeys(list(d1.keys()) + list(d2.keys())))

        # 识别电机组 key（含 "动作参数" 子 dict）
        motor_keys = [
            k for k in all_keys
            if isinstance(d1.get(k) or d2.get(k), dict)
            and "动作参数" in (d1.get(k) or d2.get(k) or {})
            and k not in ("ADP",)
        ]

        # ── 1. 系统信息 ──────────────────────────────────────────────────
        g = self._insert_group("", "系统信息")
        g_diff = False
        for k in self.SYSTEM_KEYS:
            if k in all_keys:
                g_diff |= self._insert_leaf(g, k, d1.get(k), d2.get(k))
        self._mark_group_diff(g, g_diff)
        self._diff_count += sum(1 for _, d in self._leaf_iids[-len(self.SYSTEM_KEYS):] if d)

        # ── 2. 主板参数 ──────────────────────────────────────────────────
        mb1, mb2 = d1.get("主板参数", {}), d2.get("主板参数", {})
        if mb1 or mb2:
            g = self._insert_group("", "主板参数")
            g_diff = False
            for k in dict.fromkeys(list(mb1.keys()) + list(mb2.keys())):
                g_diff |= self._insert_leaf(g, k, mb1.get(k), mb2.get(k))
            self._mark_group_diff(g, g_diff)

        # ── 3. ADP参数 ───────────────────────────────────────────────────
        adp1 = (d1.get("ADP") or {}).get("动作参数", [])
        adp2 = (d2.get("ADP") or {}).get("动作参数", [])
        if adp1 or adp2:
            g = self._insert_group("", "ADP参数")
            g_diff = False
            n = max(len(adp1), len(adp2))
            for i in range(n):
                key = f"ADP_{i}"
                named = key in self._action_names
                v1 = adp1[i] if i < len(adp1) else None
                v2 = adp2[i] if i < len(adp2) else None
                # 跳过：截图未命名 且 两侧均为 0 / None
                if not named and (v1 or 0) == 0 and (v2 or 0) == 0:
                    continue
                name = self._action_names.get(key, f"参数[{i}]")
                g_diff |= self._insert_leaf(g, name, v1, v2)
            self._mark_group_diff(g, g_diff)

        # ── 4. 电机参数组（除动作参数）────────────────────────────────────
        g_motor = self._insert_group("", "电机参数")
        g_motor_diff = False
        for mkey in motor_keys:
            m1, m2 = (d1.get(mkey) or {}), (d2.get(mkey) or {})
            sub_keys = [k for k in dict.fromkeys(list(m1.keys()) + list(m2.keys()))
                        if k != "动作参数"]
            mg = self._insert_group(g_motor, mkey, open_=False)
            mg_diff = False
            for k in sub_keys:
                mg_diff |= self._insert_leaf(mg, k, m1.get(k), m2.get(k))
            self._mark_group_diff(mg, mg_diff)
            g_motor_diff |= mg_diff
        self._mark_group_diff(g_motor, g_motor_diff)

        # ── 5. 坐标（坐标定位点 + 各电机动作参数）────────────────────────
        g_coord = self._insert_group("", "坐标")
        g_coord_diff = False

        # 5a. 坐标定位点
        coord1, coord2 = (d1.get("坐标") or {}), (d2.get("坐标") or {})
        if coord1 or coord2:
            cg = self._insert_group(g_coord, "坐标定位点")
            cg_diff = False
            all_points = dict.fromkeys(list(coord1.keys()) + list(coord2.keys()))
            for pt in all_points:
                xyz1 = coord1.get(pt, [None, None, None])
                xyz2 = coord2.get(pt, [None, None, None])
                for i, axis in enumerate(["X", "Y", "Z"]):
                    v1 = xyz1[i] if xyz1 and i < len(xyz1) else None
                    v2 = xyz2[i] if xyz2 and i < len(xyz2) else None
                    cg_diff |= self._insert_leaf(cg, f"{pt}  {axis}", v1, v2)
            self._mark_group_diff(cg, cg_diff)
            g_coord_diff |= cg_diff

        # 5b. 各电机动作参数
        for mkey in motor_keys:
            m1, m2 = (d1.get(mkey) or {}), (d2.get(mkey) or {})
            arr1 = m1.get("动作参数", [])
            arr2 = m2.get("动作参数", [])
            mg = self._insert_group(g_coord, f"{mkey}  动作参数", open_=False)
            mg_diff = False
            n = max(len(arr1), len(arr2), 20)
            for i in range(n):
                key = f"{mkey}_{i}"
                named = key in self._action_names
                v1 = arr1[i] if i < len(arr1) else None
                v2 = arr2[i] if i < len(arr2) else None
                # 跳过：截图未命名 且 两侧均为 0 / None
                if not named and (v1 or 0) == 0 and (v2 or 0) == 0:
                    continue
                name = self._action_names.get(key, f"参数[{i}]")
                mg_diff |= self._insert_leaf(mg, name, v1, v2)
            self._mark_group_diff(mg, mg_diff)
            g_coord_diff |= mg_diff
        self._mark_group_diff(g_coord, g_coord_diff)

        # ── 6. 温控参数 ──────────────────────────────────────────────────
        tc_keys = [k for k in all_keys if "温控" in k]
        if tc_keys:
            g_tc = self._insert_group("", "温控参数")
            g_tc_diff = False
            for k in tc_keys:
                tc1, tc2 = (d1.get(k) or {}), (d2.get(k) or {})
                tg = self._insert_group(g_tc, k)
                tg_diff = False
                for fk in dict.fromkeys(list(tc1.keys()) + list(tc2.keys())):
                    tg_diff |= self._insert_leaf(tg, fk, tc1.get(fk), tc2.get(fk))
                self._mark_group_diff(tg, tg_diff)
                g_tc_diff |= tg_diff
            self._mark_group_diff(g_tc, g_tc_diff)

        # ── 其他顶层 key（通道参数、管径参数等）────────────────────────────
        handled = set(self.SYSTEM_KEYS) | {"主板参数", "ADP", "坐标"} | set(motor_keys) | set(tc_keys)
        for k in all_keys:
            if k in handled:
                continue
            v1, v2 = d1.get(k), d2.get(k)
            if isinstance(v1 or v2, dict):
                sub1, sub2 = (v1 or {}), (v2 or {})
                g = self._insert_group("", k)
                g_diff = False
                for sk in dict.fromkeys(list(sub1.keys()) + list(sub2.keys())):
                    sv1, sv2 = sub1.get(sk), sub2.get(sk)
                    if isinstance(sv1 or sv2, list):
                        for i, (e1, e2) in enumerate(zip(
                            sv1 if isinstance(sv1, list) else [],
                            sv2 if isinstance(sv2, list) else []
                        )):
                            g_diff |= self._insert_leaf(g, f"{sk}[{i}]", e1, e2)
                    else:
                        g_diff |= self._insert_leaf(g, sk, sv1, sv2)
                self._mark_group_diff(g, g_diff)
            else:
                self._insert_leaf("", k, v1, v2)

        # 统计差异数
        self._diff_count = sum(1 for _, d in self._leaf_iids if d)
        self._diff_lbl.config(text=f"差异数：{self._diff_count}")

        # 应用过滤
        if self.diff_only.get():
            self._apply_filter()

    # ── 过滤（仅显示差异）────────────────────────────────────────────────────

    def _reattach_all(self):
        """将所有之前 detach 的节点还原到树中"""
        for iid, parent, index in reversed(self._detached_items):
            try:
                self.tree.reattach(iid, parent, index)
            except Exception:
                pass
        self._detached_items.clear()

    def _apply_filter(self):
        diff_only = self.diff_only.get()
        diff_set = {iid for iid, is_diff in self._leaf_iids if is_diff}

        # 先还原所有之前 detach 的节点（不重新 populate，避免递归）
        self._reattach_all()

        if not diff_only:
            return

        def process(iid: str):
            children = list(self.tree.get_children(iid))
            if not children:
                # 叶节点：无差异则隐藏
                if iid not in diff_set:
                    parent = self.tree.parent(iid)
                    index = self.tree.index(iid)
                    self._detached_items.append((iid, parent, index))
                    self.tree.detach(iid)
            else:
                for child in list(children):
                    process(child)
                # 分组节点：若子节点全被隐藏则隐藏自身
                if not self.tree.get_children(iid):
                    parent = self.tree.parent(iid)
                    index = self.tree.index(iid)
                    self._detached_items.append((iid, parent, index))
                    self.tree.detach(iid)

        for top in list(self.tree.get_children()):
            process(top)

    # ── 双击编辑动作参数名称 ─────────────────────────────────────────────────

    def _on_double_click(self, event):
        item = self.tree.focus()
        if not item:
            return
        col = self.tree.identify_column(event.x)
        label = self.tree.item(item, "text")
        if not label:
            return

        # ── 备注列编辑 (#3) ──────────────────────────────────────
        if col == "#3":
            # 只对叶节点（无子项）允许编辑备注
            if self.tree.get_children(item):
                return
            current = self._param_notes.get(label, "")
            new_note = simpledialog.askstring(
                "编辑备注说明", f"请输入「{label}」的备注（留空则清除）:",
                initialvalue=current, parent=self)
            if new_note is None:
                return
            if new_note.strip():
                self._param_notes[label] = new_note.strip()
            else:
                self._param_notes.pop(label, None)
            self._save_param_notes()
            self._populate()
            return

        if col != "#0":  # 只允许编辑参数名称列
            return
        # 找到匹配的 action_names key（格式：motor_i 或 ADP_i）
        # 通过父节点名称 + 当前行名称反推 key
        parent_iid = self.tree.parent(item)
        if not parent_iid:
            return
        parent_label = self.tree.item(parent_iid, "text")
        # 提取索引
        import re as _re
        m = _re.search(r'\[(\d+)\]', label)
        if not m:
            return
        idx = m.group(1)
        # 判断是 ADP 还是电机
        gp_iid = self.tree.parent(parent_iid)
        gp_label = self.tree.item(gp_iid, "text") if gp_iid else ""
        if "ADP" in parent_label:
            key = f"ADP_{idx}"
        else:
            # parent_label 形如 "加样Y轴  动作参数"
            motor = parent_label.split("动作参数")[0].strip()
            key = f"{motor}_{idx}"

        new_name = simpledialog.askstring(
            "编辑参数名称", f"请输入索引 [{idx}] 的名称（留空恢复默认）:",
            initialvalue=self._action_names.get(key, ""), parent=self)
        if new_name is None:
            return
        if new_name.strip():
            self._action_names[key] = new_name.strip()
        else:
            self._action_names.pop(key, None)
        self._save_action_names()
        self._populate()

    # ── 持久化 ───────────────────────────────────────────────────────────────

    def _load_action_names(self):
        path = os.path.join(self.app_dir, "param_action_names.json")
        if os.path.exists(path):
            import json as _json
            with open(path, encoding="utf-8") as f:
                self._action_names = _json.load(f)

    def _save_action_names(self):
        path = os.path.join(self.app_dir, "param_action_names.json")
        import json as _json
        with open(path, "w", encoding="utf-8") as f:
            _json.dump(self._action_names, f, ensure_ascii=False, indent=2)

    def _load_param_notes(self):
        path = os.path.join(self.app_dir, "param_notes.json")
        if os.path.exists(path):
            import json as _json
            with open(path, encoding="utf-8") as f:
                self._param_notes = _json.load(f)

    def _save_param_notes(self):
        path = os.path.join(self.app_dir, "param_notes.json")
        import json as _json
        with open(path, "w", encoding="utf-8") as f:
            _json.dump(self._param_notes, f, ensure_ascii=False, indent=2)

    # ── 搜索 ─────────────────────────────────────────────────────────────────

    def _do_search(self):
        query = self._search_var.get().strip().lower()
        if not query:
            self._apply_filter()   # 恢复差异过滤状态
            return
        # 先还原所有 detach
        self._reattach_all()

        def process(iid: str):
            children = list(self.tree.get_children(iid))
            if not children:
                label = self.tree.item(iid, "text").lower()
                if query not in label:
                    parent = self.tree.parent(iid)
                    index = self.tree.index(iid)
                    self._detached_items.append((iid, parent, index))
                    self.tree.detach(iid)
            else:
                for child in list(children):
                    process(child)
                if not self.tree.get_children(iid):
                    parent = self.tree.parent(iid)
                    index = self.tree.index(iid)
                    self._detached_items.append((iid, parent, index))
                    self.tree.detach(iid)
                else:
                    self.tree.item(iid, open=True)  # 自动展开有匹配子项的分组

        for top in list(self.tree.get_children()):
            process(top)

    def _clear_search(self):
        self._search_var.set("")


# ── ParamNotesDialog ─────────────────────────────────────────────────────────

class ParamNotesDialog(tk.Toplevel):
    """整机参数设置：为所有动作参数添加备注说明"""

    def __init__(self, parent, app_dir: str, param_view=None):
        super().__init__(parent)
        self.app_dir = app_dir
        self.param_view = param_view   # 回调刷新 ParamsView
        self.title("整机参数设置")
        self.geometry("720x540")
        self.configure(bg=C_BG)
        self.resizable(True, True)

        self._action_names: dict[str, str] = {}
        self._param_notes: dict[str, str] = {}
        self._load_data()
        self._build_ui()

    # ── 数据读写 ──────────────────────────────────────────────────────────────

    def _load_data(self):
        import json as _json
        for fname, attr in [("param_action_names.json", "_action_names"),
                             ("param_notes.json", "_param_notes")]:
            p = os.path.join(self.app_dir, fname)
            if os.path.exists(p):
                with open(p, encoding="utf-8") as f:
                    setattr(self, attr, _json.load(f))

    def _save_notes(self):
        import json as _json
        p = os.path.join(self.app_dir, "param_notes.json")
        with open(p, "w", encoding="utf-8") as f:
            _json.dump(self._param_notes, f, ensure_ascii=False, indent=2)
        # 同步刷新 ParamsView
        if self.param_view:
            self.param_view._load_param_notes()
            self.param_view._populate()

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self, text="双击「备注说明」列可编辑，备注将同步显示在整机参数对比列表中",
                 bg=C_BG, fg=C_TEXT2, font=("Microsoft YaHei", 8)
                 ).pack(anchor="w", padx=12, pady=(8, 2))

        tree_frame = tk.Frame(self, bg=C_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("电机 / 分类", "参数名称", "备注说明")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                  selectmode="browse")
        for col, w in zip(cols, [200, 180, 300]):
            self.tree.heading(col, text=col, anchor="w")
            self.tree.column(col, width=w, minwidth=60, anchor="w")
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_double_click)
        self._populate()

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        # 按电机分组，组内按索引排序
        groups: dict[str, list[tuple[int, str, str]]] = {}
        for key, name in self._action_names.items():
            parts = key.rsplit("_", 1)
            if len(parts) != 2:
                continue
            motor, idx_s = parts
            try:
                idx = int(idx_s)
            except ValueError:
                continue
            groups.setdefault(motor, []).append((idx, name, key))
        for motor in sorted(groups):
            for idx, name, key in sorted(groups[motor]):
                note = self._param_notes.get(name, "")
                self.tree.insert("", tk.END, iid=key, values=(motor, name, note))

    def _on_double_click(self, event):
        item = self.tree.focus()
        if not item:
            return
        col = self.tree.identify_column(event.x)
        if col != "#3":
            return
        motor, name, current_note = self.tree.item(item, "values")
        new_note = simpledialog.askstring(
            "编辑备注说明", f"请输入「{name}」的备注（留空则清除）:",
            initialvalue=current_note, parent=self)
        if new_note is None:
            return
        if new_note.strip():
            self._param_notes[name] = new_note.strip()
        else:
            self._param_notes.pop(name, None)
        self.tree.item(item, values=(motor, name, self._param_notes.get(name, "")))
        self._save_notes()


# ── StandardFlowDialog ───────────────────────────────────────────────────────

class StandardFlowDialog(tk.Toplevel):
    """展示并管理各模式标准动作流程列表。"""

    # 已知测试模式（展示顺序）
    KNOWN_MODES = ["两孔稀释", "单次稀释", "多次稀释", "一步法"]

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("标准流程管理")
        self.geometry("680x500")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self._build_ui()

    def _build_ui(self):
        # ── 顶部操作栏 ──
        bar = tk.Frame(self, bg=C_BG)
        bar.pack(fill=tk.X, padx=10, pady=8)

        self._mode_var = tk.StringVar()
        modes_with_data = [m for m in self.KNOWN_MODES
                           if m in self.app.parser.standard_sequences]
        all_modes = modes_with_data + [m for m in self.KNOWN_MODES
                                       if m not in modes_with_data]
        self._mode_combo = ttk.Combobox(bar, textvariable=self._mode_var,
                                        values=all_modes, width=12,
                                        state="readonly")
        self._mode_combo.pack(side=tk.LEFT)
        if all_modes:
            self._mode_combo.set(all_modes[0])
        self._mode_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_tree())

        tk.Button(bar, text="载入参考日志", command=self._load_ref_log,
                  bg=C_BLUE, fg=C_WHITE, relief="flat", padx=10).pack(side=tk.LEFT, padx=6)
        tk.Button(bar, text="保存", command=self._save,
                  bg=C_GREEN, fg=C_WHITE, relief="flat", padx=10).pack(side=tk.LEFT)
        tk.Button(bar, text="清除当前模式", command=self._clear_current,
                  bg=C_CARD, fg=C_TEXT, relief="flat", padx=10).pack(side=tk.LEFT, padx=6)

        self._status_var = tk.StringVar(value="")
        tk.Label(bar, textvariable=self._status_var, bg=C_BG,
                 fg=C_TEXT2, font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=8)

        # ── 动作序列树状列表 ──
        frame = tk.Frame(self, bg=C_CARD, bd=1, relief="solid")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.tree = ttk.Treeview(frame, columns=("seq", "code", "name"),
                                 show="headings", selectmode="browse")
        self.tree.heading("seq",  text="顺序", anchor="center")
        self.tree.heading("code", text="动作编号", anchor="center")
        self.tree.heading("name", text="动作名称", anchor="w")
        self.tree.column("seq",  width=60,  anchor="center", stretch=False)
        self.tree.column("code", width=100, anchor="center", stretch=False)
        self.tree.column("name", width=400, anchor="w")

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self._refresh_tree()

    def _get_action_name(self, code: str) -> str:
        p = self.app.parser
        return (p.theory_display_names.get(code)
                or p.theory_names.get(code)
                or p.action_names.get(code, ""))

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        mode = self._mode_var.get()
        seq = self.app.parser.standard_sequences.get(mode, [])
        for i, code in enumerate(seq, 1):
            name = self._get_action_name(code)
            self.tree.insert("", tk.END, values=(i, code, name))
        count = len(seq)
        self._status_var.set(f"共 {count} 个动作" if count else "（暂无标准流程）")

    def _load_ref_log(self):
        """从参考日志文件中提取样本0001的动作序列，设为当前模式的标准。"""
        mode = self._mode_var.get()
        if not mode:
            return
        path = filedialog.askopenfilename(
            title=f"选择「{mode}」参考日志文件",
            filetypes=[("文本/日志文件", "*.txt *.log"), ("所有文件", "*.*")])
        if not path:
            return

        try:
            ref_parser = LogParser()
            # 传入当前理论时间设置，以便动作名称解析一致
            ref_parser.theory_times = dict(self.app.parser.theory_times)
            ref_parser.theory_names = dict(self.app.parser.theory_names)
            ref_parser.theory_display_names = dict(self.app.parser.theory_display_names)
            ref_parser.action_names = dict(self.app.parser.action_names)
            ref_parser.load_file(path)
        except Exception as e:
            messagebox.showerror("解析失败", str(e), parent=self)
            return

        ref_actions = ref_parser.actions.get("0001", [])
        if not ref_actions:
            messagebox.showwarning("未找到数据",
                                   "参考日志中未找到样本0001的动作记录。", parent=self)
            return

        seen: set[str] = set()
        sequence: list[str] = []
        for a in ref_actions:
            if a.level1 not in seen:
                seen.add(a.level1)
                sequence.append(a.level1)

        self.app.parser.standard_sequences[mode] = sequence
        self._refresh_tree()
        messagebox.showinfo("载入成功",
                            f"已从参考日志提取 {len(sequence)} 个动作，\n"
                            f"设为「{mode}」标准流程。\n点击「保存」持久化。",
                            parent=self)

    def _clear_current(self):
        mode = self._mode_var.get()
        if mode and mode in self.app.parser.standard_sequences:
            if messagebox.askyesno("确认", f"确定清除「{mode}」的标准流程吗？", parent=self):
                del self.app.parser.standard_sequences[mode]
                self._refresh_tree()

    def _save(self):
        app_dir = _app_dir()
        path = os.path.join(app_dir, "standard_sequences.json")
        self.app.parser.save_standard_sequences(path)
        self._status_var.set("已保存")


# ── Main Application ─────────────────────────────────────────────────────────

class FA120App:
    def __init__(self, root):
        self.root = root
        self.root.title("FA120 日志解析软件 v1.0")
        self.root.geometry("1400x800")
        self.root.minsize(1000, 600)
        self.root.configure(bg=C_BG)

        self.parser = LogParser()
        self._loading = False
        self._load_executor = None
        self._load_future = None
        self._self_check_details = {}
        self._syncing_sample_selection = False

        # 尝试加载预设文件
        self._load_presets()

        self._apply_style()
        self._build_ui()

    def _load_presets(self):
        """加载预设文件: 运动部件名称 + 理论时间表 + 一级动作名称"""
        app_dir = _app_dir()

        # 运动部件名称 - 优先同目录，再桌面
        motor_paths = [
            os.path.join(app_dir, "运动部件名称.txt"),
            os.path.expanduser("~/Desktop/运动部件名称.txt"),
        ]
        for p in motor_paths:
            if os.path.exists(p):
                self.parser.load_motor_names(p)
                break

        # 理论时间表 — 优先加载 JSON（含用户历史编辑），不存在则从 txt 初始化
        theory_json = os.path.join(app_dir, "theory_data.json")
        if os.path.exists(theory_json):
            import json
            with open(theory_json, encoding="utf-8") as f:
                data = json.load(f)
            self.parser.theory_names = data.get("theory_names", {})
            self.parser.theory_display_names = data.get("theory_display_names", {})
            self.parser.theory_times = {k: int(v) for k, v in data.get("theory_times", {}).items()}
        else:
            theory_paths = [
                os.path.join(app_dir, "理论时间表.txt"),
                os.path.join(app_dir, "理论时间表举例.txt"),
                os.path.expanduser("~/Desktop/理论时间表举例.txt"),
                os.path.expanduser("~/Desktop/理论时间表.txt"),
            ]
            for p in theory_paths:
                if os.path.exists(p):
                    self.parser.load_theory_time(p)
                    break

        # 电机别名
        alias_path = os.path.join(app_dir, "motor_aliases.json")
        if os.path.exists(alias_path):
            import json
            with open(alias_path, encoding="utf-8") as f:
                self.parser.motor_aliases = json.load(f)

        # 历史部件列表（无日志时也能显示完整电机名称设置）
        comp_path = os.path.join(app_dir, "known_components.json")
        if os.path.exists(comp_path):
            import json
            with open(comp_path, encoding="utf-8") as f:
                known = json.load(f)
            existing = set(self.parser.all_components)
            for c in known:
                if c not in existing:
                    self.parser.all_components.append(c)
                    existing.add(c)

        # 各模式标准动作序列
        std_seq_path = os.path.join(app_dir, "standard_sequences.json")
        self.parser.load_standard_sequences(std_seq_path)

        # 一级动作名称（例如桌面的 FA120动作日志帧头）
        action_name_paths = [
            os.path.join(app_dir, "FA120动作日志帧头.txt"),
            os.path.join(app_dir, "FA120动作日志帧头"),
            os.path.expanduser("~/Desktop/FA120动作日志帧头.txt"),
            os.path.expanduser("~/Desktop/FA120动作日志帧头"),
        ]
        for p in action_name_paths:
            if os.path.exists(p):
                self.parser.load_action_names(p)
                break

    def _apply_style(self):
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure("TCombobox",
                    fieldbackground=C_WHITE, background=C_WHITE,
                    foreground=C_TEXT, bordercolor=C_BORDER,
                    arrowcolor=C_TEXT2, relief="flat")
        s.map("TCombobox",
              fieldbackground=[("readonly", C_WHITE)],
              bordercolor=[("focus", C_BLUE)])
        s.configure("TCheckbutton",
                    background=C_BG2, foreground=C_TEXT,
                    font=("Microsoft YaHei", 8))
        s.map("TCheckbutton", background=[("active", C_BG2)])
        s.configure("Treeview",
                    rowheight=24, font=("Microsoft YaHei", 9),
                    background=C_WHITE, fieldbackground=C_WHITE,
                    foreground=C_TEXT, bordercolor=C_BORDER)
        s.configure("Treeview.Heading",
                    font=("Microsoft YaHei", 9, "bold"),
                    background=C_BLUE, foreground=C_WHITE,
                    relief="flat", borderwidth=0)
        s.map("Treeview.Heading",
              background=[("active", C_BLUE_HV)],
              foreground=[("active", C_WHITE)])
        s.configure("TNotebook", background=C_BG, borderwidth=0)
        s.configure("TNotebook.Tab",
                    font=("Microsoft YaHei", 10),
                    padding=(18, 10, 18, 10),
                    background="#d9d9d9",
                    foreground=C_TEXT,
                    borderwidth=1,
                    relief="flat")
        s.map("TNotebook.Tab",
              padding=[("selected", (18, 10, 18, 10)),
                       ("!selected", (18, 10, 18, 10))],
              background=[("selected", C_WHITE)],
              foreground=[("selected", C_TEXT)],
              relief=[("selected", "flat"), ("!selected", "flat")])

    def _build_ui(self):
        toolbar = tk.Frame(self.root, bg=C_BG2, pady=5)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        self.load_btn = tk.Button(toolbar, text="载入日志", command=self._load_log,
                                  bg=C_BLUE, fg=C_WHITE, font=("Microsoft YaHei", 9),
                                  relief="flat", padx=12, pady=3, cursor="hand2")
        self.load_btn.pack(side=tk.LEFT, padx=8)
        tk.Button(toolbar, text="动作设置", command=self._open_theory_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 9),
                  relief="flat", padx=12, pady=3, bd=1, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(toolbar, text="电机名称设置", command=self._open_action_name_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 9),
                  relief="flat", padx=12, pady=3, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(toolbar, text="整机参数设置", command=self._open_param_notes_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 9),
                  relief="flat", padx=12, pady=3, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(toolbar, text="标准流程", command=self._open_standard_flow_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 9),
                  relief="flat", padx=12, pady=3, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)

        # 文件名显示
        self.file_label = tk.Label(toolbar, text="未载入文件", bg=C_BG2,
                                   fg=C_TEXT2, font=("Microsoft YaHei", 8))
        self.file_label.pack(side=tk.RIGHT, padx=10)

        self.main_notebook = ttk.Notebook(self.root)
        self.main_notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self.instrument_tab = tk.Frame(self.main_notebook, bg=C_CARD, bd=1, relief="solid")
        self.sample_tab = tk.Frame(self.main_notebook, bg=C_CARD, bd=1, relief="solid")
        self.alarm_tab = tk.Frame(self.main_notebook, bg=C_CARD, bd=1, relief="solid")
        self.params_tab = tk.Frame(self.main_notebook, bg=C_CARD, bd=1, relief="solid")

        self.main_notebook.add(self.instrument_tab, text="仪器信息")
        self.main_notebook.add(self.sample_tab, text="样本信息")
        self.main_notebook.add(self.alarm_tab, text="异常信息")
        self.main_notebook.add(self.params_tab, text="整机参数")

        app_dir = _app_dir()
        self._param_view = ParamsView(self.params_tab, app_dir)
        self._param_view.pack(fill=tk.BOTH, expand=True)

        self._build_instrument_tab()
        self._build_sample_tab()
        self._build_alarm_tab()

    def _create_tree(self, parent, columns, widths=None, height=None):
        frame = tk.Frame(parent, bg=C_CARD)
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse",
                            height=height)
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        for idx, col in enumerate(columns):
            width = widths[idx] if widths and idx < len(widths) else 120
            tree.heading(col, text=col, anchor="center")
            tree.column(col, width=width, minwidth=60, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        return frame, tree

    def _build_instrument_tab(self):
        summary_frame = tk.LabelFrame(self.instrument_tab, text="关于本机",
                                      bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 10, "bold"),
                                      padx=8, pady=6)
        summary_frame.pack(fill=tk.X, padx=10, pady=(6, 4))
        self.instrument_value_labels = {}
        field_columns = [
            [
                ("仪器序列号", "device_serial"),
                ("日志时间", "log_time"),
                ("4G CCID", "ccid"),
                ("信号强度", "signal_strength"),
                ("用户程序版本", "user_program_version"),
            ],
            [
                ("中位机版本", "mid_version"),
                ("MCU1版本", "mcu1_version"),
                ("MCU2版本", "mcu2_version"),
                ("MCU3版本", "mcu3_version"),
                ("温控版本", "temp_control_version"),
            ],
            [
                ("累计申请次数", "request_count"),
                ("累计检测次数", "detect_count"),
                ("累计开盖次数", "open_cap_count"),
                ("日志编排数", "current_arrangement_count"),
            ],
        ]
        for col, fields in enumerate(field_columns):
            base_col = col * 2
            for row, (label_text, key) in enumerate(fields):
                tk.Label(summary_frame, text=f"{label_text}:", bg=C_CARD, fg=C_TEXT,
                         font=("Microsoft YaHei", 9, "bold")).grid(
                             row=row, column=base_col, sticky="w", padx=(0, 6), pady=2)
                value_label = tk.Label(summary_frame, text="-", bg=C_CARD, fg=C_TEXT2,
                                       font=("Microsoft YaHei", 9), anchor="w")
                value_label.grid(row=row, column=base_col + 1, sticky="w", padx=(0, 14), pady=2)
                self.instrument_value_labels[key] = value_label
        for col in range(6):
            summary_frame.grid_columnconfigure(col, weight=1 if col % 2 else 0)

        selfcheck_frame = tk.LabelFrame(self.instrument_tab, text="自检信息",
                                        bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 10, "bold"),
                                        padx=6, pady=6)
        selfcheck_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 4))
        selfcheck_split = tk.Frame(selfcheck_frame, bg=C_CARD)
        selfcheck_split.pack(fill=tk.BOTH, expand=True)
        selfcheck_split.grid_columnconfigure(0, weight=1, uniform="instrument_halves")
        selfcheck_split.grid_columnconfigure(1, weight=1, uniform="instrument_halves")
        selfcheck_split.grid_rowconfigure(0, weight=1)

        self.self_check_only_issues_var = tk.BooleanVar(value=False)
        self.self_check_filter_cb = tk.Checkbutton(
            selfcheck_frame,
            text="仅显示异常",
            variable=self.self_check_only_issues_var,
            command=self._populate_instrument_tab,
            bg=C_CARD,
            fg=C_TEXT,
            activebackground=C_CARD,
            activeforeground=C_TEXT,
            selectcolor=C_CARD,
            font=("Microsoft YaHei", 9),
        )
        self.self_check_filter_cb.place(relx=1.0, x=-10, y=2, anchor="ne")

        selfcheck_left = tk.Frame(selfcheck_split, bg=C_CARD)
        selfcheck_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        tree_frame, self.self_check_tree = self._create_tree(
            selfcheck_left,
            ("自检时间", "MCU编号", "状态"),
            widths=[90, 80, 100],
            height=5,
        )
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.self_check_tree.bind("<<TreeviewSelect>>", self._on_self_check_select)
        self.self_check_tree.tag_configure("error_status", foreground=C_RED)
        self.self_check_tree.tag_configure("warn_status", foreground="#ef6c00")

        detail_frame = tk.Frame(selfcheck_split, bg=C_CARD)
        detail_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        tk.Label(detail_frame, text="错误详情", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 9, "bold")).pack(anchor="w", pady=(0, 4))

        text_frame = tk.Frame(detail_frame, bg=C_CARD)
        text_frame.pack(fill=tk.BOTH, expand=True)
        self.self_check_detail = tk.Text(
            text_frame,
            height=6,
            wrap="word",
            bg=C_WHITE,
            fg=C_TEXT,
            font=("Microsoft YaHei", 9),
            relief="solid",
            bd=1,
        )
        detail_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.self_check_detail.yview)
        self.self_check_detail.configure(yscrollcommand=detail_scroll.set)
        self.self_check_detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        detail_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._set_self_check_detail("请选择一条自检记录查看详情")

        action_frame = tk.LabelFrame(self.instrument_tab, text="用户动作",
                                     bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 10, "bold"),
                                     padx=6, pady=6)
        action_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))
        action_split = tk.Frame(action_frame, bg=C_CARD)
        action_split.pack(fill=tk.BOTH, expand=True)
        action_split.grid_columnconfigure(0, weight=1, uniform="instrument_halves")
        action_split.grid_columnconfigure(1, weight=1, uniform="instrument_halves")
        action_split.grid_rowconfigure(0, weight=1)

        action_left = tk.Frame(action_split, bg=C_CARD)
        action_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        tree_frame, self.user_action_tree = self._create_tree(
            action_left,
            ("动作时间", "动作", "详情"),
            widths=[90, 110, 220],
            height=8,
        )
        tree_frame.pack(fill=tk.BOTH, expand=True)

        action_placeholder = tk.Frame(action_split, bg=C_CARD, bd=1, relief="solid")
        action_placeholder.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        tk.Label(action_placeholder, text="预留区域", bg=C_CARD, fg=C_TEXT2,
                 font=("Microsoft YaHei", 10)).pack(expand=True)

    def _build_sample_tab(self):
        sample_list_frame = tk.LabelFrame(self.sample_tab, text="样本编排信息",
                                          bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 10, "bold"),
                                          padx=8, pady=8)
        sample_list_frame.pack(fill=tk.X, padx=10, pady=(6, 8))

        # ── 统计栏 + 筛选复选框 ──
        stat_frame = tk.Frame(sample_list_frame, bg=C_CARD)
        stat_frame.pack(fill=tk.X, pady=(0, 4))

        self._stat_total_var  = tk.StringVar(value="总计: 0")
        self._stat_done_var   = tk.StringVar(value="测试完成: 0")
        self._stat_error_var  = tk.StringVar(value="异常: 0")
        self._stat_unknown_var = tk.StringVar(value="未知: 0")

        tk.Label(stat_frame, textvariable=self._stat_total_var,
                 bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(stat_frame, textvariable=self._stat_done_var,
                 bg=C_CARD, fg="#1a7f3c", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(stat_frame, textvariable=self._stat_error_var,
                 bg=C_CARD, fg=C_RED, font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(stat_frame, textvariable=self._stat_unknown_var,
                 bg=C_CARD, fg=C_TEXT2, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(stat_frame, text="显示：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self._filter_done_var    = tk.BooleanVar(value=True)
        self._filter_error_var   = tk.BooleanVar(value=True)
        self._filter_unknown_var = tk.BooleanVar(value=True)
        for text, var, fg in [
            ("测试完成", self._filter_done_var,    "#1a7f3c"),
            ("异常",     self._filter_error_var,   C_RED),
            ("未知",     self._filter_unknown_var, C_TEXT2),
        ]:
            tk.Checkbutton(stat_frame, text=text, variable=var,
                           command=self._apply_sample_filter,
                           bg=C_CARD, fg=fg, selectcolor=C_CARD,
                           font=("Microsoft YaHei", 9),
                           activebackground=C_CARD).pack(side=tk.LEFT, padx=4)

        tree_frame, self.sample_tree = self._create_tree(
            sample_list_frame,
            ("编排时间", "流水号", "样本ID", "样本类型", "样本位置", "测试数", "开盖", "摇匀", "项目", "项目缩写", "项目模式", "样本状态", "完成时间", "浓度", "测试值"),
            widths=[80, 60, 130, 60, 60, 45, 45, 45, 100, 80, 80, 80, 80, 70, 90],
            height=7,
        )
        tree_frame.pack(fill=tk.X, expand=True)
        self.sample_tree.bind("<<TreeviewSelect>>", self._on_sample_tree_select)

        ttk.Separator(self.sample_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        detail_frame = tk.LabelFrame(self.sample_tab, text="样本动作详情",
                                     bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 10, "bold"),
                                     padx=8, pady=8)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        tab_frame = tk.Frame(detail_frame, bg=C_BG2)
        tab_frame.pack(fill=tk.X, pady=(0, 6))

        self._view_var = tk.StringVar(value="table")
        tk.Radiobutton(tab_frame, text="表格视图", variable=self._view_var,
                       value="table", command=self._switch_view,
                       bg=C_BG2, fg=C_TEXT, font=("Microsoft YaHei", 9),
                       selectcolor=C_BG2, indicatoron=0, padx=15, pady=3,
                       relief="flat").pack(side=tk.LEFT, padx=2, pady=2)
        tk.Radiobutton(tab_frame, text="时间轴视图", variable=self._view_var,
                       value="timeline", command=self._switch_view,
                       bg=C_BG2, fg=C_TEXT, font=("Microsoft YaHei", 9),
                       selectcolor=C_BG2, indicatoron=0, padx=15, pady=3,
                       relief="flat").pack(side=tk.LEFT, padx=2, pady=2)

        self.view_container = tk.Frame(detail_frame, bg=C_CARD)
        self.view_container.pack(fill=tk.BOTH, expand=True)
        self.timeline = TimelineCanvas(self.view_container, self)
        self.table_view = TableView(self.view_container, self)
        self.table_view.pack(fill=tk.BOTH, expand=True)

    def _build_alarm_tab(self):
        tk.Label(self.alarm_tab, text="异常信息", bg=C_CARD, fg=C_RED,
                 font=("Microsoft YaHei", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
        alarm_tree_frame, self.alarm_tree = self._create_tree(
            self.alarm_tab,
            ("时间", "错误编号", "样本", "动作", "报警内容", "详情"),
            widths=[90, 100, 80, 80, 180, 520],
        )
        alarm_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.alarm_tree.bind("<<TreeviewSelect>>", self._on_alarm_select)

    def _switch_view(self):
        view = self._view_var.get()
        if view == "timeline":
            self.table_view.pack_forget()
            self.timeline.pack(fill=tk.BOTH, expand=True)
        else:
            self.timeline.pack_forget()
            self.table_view.pack(fill=tk.BOTH, expand=True)

    def _load_log(self):
        if self._loading:
            return
        path = filedialog.askopenfilename(
            title="载入日志文件",
            filetypes=[("文本文件", "*.txt"), ("日志文件", "*.log"), ("所有文件", "*.*")])
        if not path:
            return
        self._loading = True
        self.load_btn.config(state="disabled")
        self.file_label.config(text=f"正在载入: {os.path.basename(path)}")
        if self._load_executor is None:
            self._load_executor = concurrent.futures.ProcessPoolExecutor(max_workers=1)
        self._load_future_path = path
        self._load_future = self._load_executor.submit(parse_log_job, path, self._get_preset_state())
        self.root.after(100, self._poll_load_result)

    def _get_preset_state(self):
        return {
            "theory_times": dict(self.parser.theory_times),
            "theory_names": dict(self.parser.theory_names),
            "theory_display_names": dict(self.parser.theory_display_names),
            "action_names": dict(self.parser.action_names),
            "motor_names": list(self.parser.motor_names),
            "motor_aliases": dict(self.parser.motor_aliases),
            "standard_sequences": {k: list(v) for k, v in self.parser.standard_sequences.items()},
        }

    def _poll_load_result(self):
        if not self._loading or self._load_future is None:
            return
        if not self._load_future.done():
            self.root.after(100, self._poll_load_result)
            return
        path = self._load_future_path if hasattr(self, "_load_future_path") else ""
        try:
            parser = self._load_future.result()
        except Exception as e:
            self._on_load_failed(path, e)
            return
        self._on_log_loaded(path, parser)

    def _on_load_failed(self, path: str, error: Exception):
        self._loading = False
        self._load_future = None
        self.load_btn.config(state="normal")
        self.file_label.config(text=f"载入失败: {os.path.basename(path)}")
        messagebox.showerror("错误", f"解析日志失败:\n{error}")

    def _on_log_loaded(self, path: str, parser: LogParser):
        self.parser = parser
        self._loading = False
        self._load_future = None
        self.load_btn.config(state="normal")

        self.file_label.config(text=os.path.basename(path))

        serials = sorted(self.parser.samples.keys())
        self._populate_instrument_tab()
        self._populate_sample_tree()
        self._populate_alarm_tree()

        if serials:
            self.timeline.set_data([], [], self.parser.motor_names)
            self.table_view.set_data([], [], self.parser.theory_times)
        else:
            self._select_sample("")

    def _on_sample_tree_select(self, event):
        if self._syncing_sample_selection:
            return
        sel = self.sample_tree.selection()
        if not sel:
            return
        self._select_sample(sel[0])

    def _select_sample(self, serial: str):
        if not serial:
            self.timeline.set_data([], [], self.parser.motor_names)
            self.table_view.set_data([], [], self.parser.theory_times)
            return

        if serial in self.sample_tree.get_children():
            current_sel = self.sample_tree.selection()
            if current_sel != (serial,):
                self._syncing_sample_selection = True
                try:
                    self.sample_tree.selection_set(serial)
                    self.sample_tree.focus(serial)
                    self.sample_tree.see(serial)
                finally:
                    self._syncing_sample_selection = False

        actions = self.parser.actions.get(serial, [])
        sample_alarms = [a for a in self.parser.alarms if a.sample_num == serial]
        self.timeline.set_data(actions, sample_alarms, self.parser.motor_names,
                               self.parser.system_actions)
        sample = self.parser.samples.get(serial)
        missing = sample.missing_actions if sample else []
        # 使用该样本测试模式对应的标准序列（而非全局 standard_action_sequence）
        mode_standard = self.parser.standard_sequences.get(
            sample.mode if sample else "", []
        ) or self.parser.standard_action_sequence
        self.table_view.set_data(actions, sample_alarms, self.parser.theory_times,
                                 standard_sequence=mode_standard,
                                 missing_actions=missing)

    def _on_alarm_select(self, event):
        sel = self.alarm_tree.selection()
        if not sel:
            return
        vals = self.alarm_tree.item(sel[0], "values")
        if not vals:
            return
        sample_serial = vals[2]  # 样本流水号
        action_code = vals[3]    # 动作编号
        alarm_time = vals[0]     # 报警时间

        serials = list(self.parser.samples.keys())
        if sample_serial in serials:
            self._select_sample(sample_serial)
            t_ms = time_to_ms(alarm_time + ".000") if '.' not in alarm_time else time_to_ms(alarm_time)
            self.timeline.scroll_to_time(t_ms)
            self.table_view.highlight_action(action_code)

    def highlight_alarm(self, alarm: Alarm):
        """从时间轴点击报警后，高亮异常信息列表"""
        for item in self.alarm_tree.get_children():
            vals = self.alarm_tree.item(item, "values")
            if vals and vals[1] == alarm.error_code and vals[2] == alarm.sample_num:
                self.alarm_tree.selection_set(item)
                self.alarm_tree.see(item)
                break

    def _set_self_check_detail(self, text: str):
        self.self_check_detail.config(state="normal")
        self.self_check_detail.delete("1.0", tk.END)
        self.self_check_detail.insert("1.0", text)
        self.self_check_detail.config(state="disabled")

    def _on_self_check_select(self, event):
        sel = self.self_check_tree.selection()
        if not sel:
            self._set_self_check_detail("请选择一条自检记录查看详情")
            return
        item_id = sel[0]
        values = self.self_check_tree.item(item_id, "values")
        if not values:
            self._set_self_check_detail("请选择一条自检记录查看详情")
            return

        status = values[2]
        full_error = self._self_check_details.get(item_id, "")
        if status == "完成" or not full_error:
            self._set_self_check_detail("无错误")
        else:
            self._set_self_check_detail(full_error)

    def _populate_instrument_tab(self):
        info = self.parser.instrument_info
        for key, label in self.instrument_value_labels.items():
            value = getattr(info, key, "")
            label.config(text=str(value) if value not in ("", None) else "-")

        self.self_check_tree.delete(*self.self_check_tree.get_children())
        self._self_check_details = {}
        for idx, record in enumerate(self.parser.self_checks):
            if self.self_check_only_issues_var.get() and record.status == "完成":
                continue
            item_id = f"selfcheck-{idx}"
            tags = ()
            if record.status == "错误":
                tags = ("error_status",)
            elif record.status == "异常后完成":
                tags = ("warn_status",)
            self.self_check_tree.insert("", tk.END, iid=item_id, values=(
                record.check_time,
                record.mcu,
                record.status,
            ), tags=tags)
            self._self_check_details[item_id] = record.error_info or ""
        self._set_self_check_detail("请选择一条自检记录查看详情")

        self.user_action_tree.delete(*self.user_action_tree.get_children())
        for idx, record in enumerate(self.parser.user_actions):
            self.user_action_tree.insert("", tk.END, iid=f"useraction-{idx}", values=(
                record.action_time,
                record.action_type,
                record.detail,
            ))

    def _populate_sample_tree(self):
        # 统计各状态数量
        all_samples = list(self.parser.samples.values())
        n_total   = len(all_samples)
        n_done    = sum(1 for s in all_samples if s.status == "测试完成")
        n_error   = sum(1 for s in all_samples if s.status == "异常")
        n_unknown = n_total - n_done - n_error

        self._stat_total_var.set(f"总计: {n_total}")
        self._stat_done_var.set(f"测试完成: {n_done}")
        self._stat_error_var.set(f"异常: {n_error}")
        self._stat_unknown_var.set(f"未知: {n_unknown}")

        # 配置行颜色标签
        self.sample_tree.tag_configure("status_done",    background="#e8f8ee")
        self.sample_tree.tag_configure("status_error",   background="#ffd6d6")
        self.sample_tree.tag_configure("status_unknown",  background="#f5f5f5")

        show_done    = self._filter_done_var.get()
        show_error   = self._filter_error_var.get()
        show_unknown = self._filter_unknown_var.get()

        self.sample_tree.delete(*self.sample_tree.get_children())
        for serial in sorted(self.parser.samples.keys()):
            sample = self.parser.samples[serial]

            # 筛选过滤
            if sample.status == "测试完成" and not show_done:
                continue
            if sample.status == "异常" and not show_error:
                continue
            if sample.status not in ("测试完成", "异常") and not show_unknown:
                continue

            item_text = "、".join(sample.test_items[:2])
            if sample.status == "测试完成":
                row_tag = ("status_done",)
            elif sample.status == "异常":
                row_tag = ("status_error",)
            else:
                row_tag = ("status_unknown",)

            self.sample_tree.insert("", tk.END, iid=serial, tags=row_tag, values=(
                sample.arrange_time,
                sample.serial,
                sample.sample_id,
                sample.sample_type,
                self.parser._plus_one_rack_pos(sample.rack_pos),
                sample.test_count,
                "是" if sample.cap_open else "否",
                "是" if sample.shake else "否",
                item_text,
                sample.project_abbr,
                sample.mode,
                sample.status,
                sample.finish_time,
                sample.concentration,
                sample.measure_value,
            ))

    def _apply_sample_filter(self):
        """筛选复选框变化时刷新样本列表（保留当前选中样本）。"""
        sel = self.sample_tree.selection()
        prev_serial = sel[0] if sel else ""
        self._populate_sample_tree()
        if prev_serial and prev_serial in self.sample_tree.get_children():
            self.sample_tree.selection_set(prev_serial)
            self.sample_tree.see(prev_serial)

    def _populate_alarm_tree(self):
        self.alarm_tree.delete(*self.alarm_tree.get_children())
        for idx, alarm in enumerate(self.parser.alarms):
            self.alarm_tree.insert("", tk.END, iid=f"alarm-{idx}", values=(
                alarm.time_str,
                alarm.error_code,
                alarm.sample_num,
                alarm.action_code,
                alarm.content,
                alarm.detail,
            ))

    def _open_params_tab(self):
        self.main_notebook.select(self.params_tab)

    def _open_param_notes_dialog(self):
        app_dir = _app_dir()
        ParamNotesDialog(self.root, app_dir, param_view=self._param_view)

    def _open_standard_flow_dialog(self):
        StandardFlowDialog(self.root, self)

    def _open_theory_dialog(self):
        app_dir = _app_dir()
        save_path = os.path.join(app_dir, "theory_data.json")
        TheoryTimeDialog(self.root, self.parser, on_change=self._refresh_views, save_path=save_path)

    def _open_action_name_dialog(self):
        app_dir = _app_dir()
        save_path = os.path.join(app_dir, "motor_aliases.json")
        # 打开前先把当前 all_components 持久化，下次无日志也能显示完整列表
        self._save_known_components(app_dir)
        MotorNamesDialog(self.root, self.parser, on_change=self._refresh_views, save_path=save_path)

    def _save_known_components(self, app_dir: str):
        import json
        path = os.path.join(app_dir, "known_components.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.parser.all_components, f, ensure_ascii=False, indent=2)

    def _refresh_views(self):
        self._populate_instrument_tab()
        self._populate_sample_tree()
        self._populate_alarm_tree()
        sel = self.sample_tree.selection()
        serial = sel[0] if sel else ""
        if serial:
            self._select_sample(serial)


# ── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = FA120App(root)
    root.mainloop()
