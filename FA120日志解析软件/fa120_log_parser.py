"""
FA120 日志解析软件
解析全自动荧光免疫分析仪的日志，用于异常故障排查和时序检查
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re
import os
import sys
import json
import concurrent.futures
import multiprocessing
from dataclasses import dataclass, field
import tkinter.font as tkfont


def _app_dir() -> str:
    """返回配置文件所在目录：打包为 EXE 时取 EXE 旁边的目录，否则取脚本目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

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
    test_results: list = None     # 每项目测试结果: [{project_abbr, concentration, measure_value, finish_time}]

    def __post_init__(self):
        if self.missing_actions is None:
            self.missing_actions = []
        if self.test_results is None:
            self.test_results = []


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
    source_line: int = 0   # 日志原文行号（1-based）


@dataclass
class Alarm:
    time_str: str        # 报警时间
    error_code: str      # 错误编号 "G-INJ-012"
    display_code: str    # 原始报警编号（用于列表显示）
    sample_num: str      # 对应样本 "0005"
    action_code: str     # 对应动作 "E01"
    content: str         # 报警内容
    detail: str          # 详情
    source_line: int = 0


@dataclass
class InstrumentInfo:
    device_serial: str = ""
    log_time: str = ""
    log_range: str = ""
    user_program_version: str = ""
    request_count: int = 0
    detect_count: int = 0
    open_cap_count: int = 0
    current_arrangement_count: int = 0
    ccid: str = ""
    signal_strength: str = ""
    mid_version: str = ""
    mcu0_version: str = ""
    mcu1_version: str = ""
    mcu2_version: str = ""
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
    source_line: int = 0
    related_lines: list[int] = field(default_factory=list)


# ── Log Parser ───────────────────────────────────────────────────────────────

class LogParser:
    # 样本申请
    RE_SAMPLE = re.compile(
        r'"#(\d+)-(\d{4})\s*"\s*"申请：样本架(\d+-\d+)，样本ID:(\w*),\s*测试数(\d+)，'
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
        r'(\d{2}:\d{2}:\d{2}).*?报警信息\s*"([^"]+)"\s*'
        r'"([^"]*)"\s*"([^"]*)"\s*"([^"]*)"'
    )

    RE_TIME_PREFIX = re.compile(r'^(\d{2}:\d{2}:\d{2})')
    RE_SELF_CHECK = re.compile(r'Msg:\s*"([\d:.]+)"\s*"MCU([123])自检"')
    RE_SELF_CHECK_FAIL = re.compile(r'报警信息\s*"(G-OTH-20[234])')
    RE_USER_LOGIN = re.compile(r'user:\s*"([^"]*)"', re.IGNORECASE)
    RE_DEVICE_SERIAL = re.compile(r'仪器序列号.*?"([^"]+)"')
    RE_CCID = re.compile(r'iccid\s*"([^"]+)"', re.IGNORECASE)
    RE_SIGNAL = re.compile(r'信号值\s+([^\r\n"]+)')
    RE_MID_VERSION = re.compile(r'MCU3Mid version.*?"(v?1\.0[^"]*)"', re.IGNORECASE)
    RE_FW_VERSION = re.compile(r'MCU([0124])\s+SN:\s*".*?"\s*"([^"]+)"', re.IGNORECASE)
    RE_STATS = re.compile(r'申请样本次数[:：](\d+)\s+检测次数(\d+)\s+开盖次数(\d+)')
    RE_LOAD_PROJECT = re.compile(r'导入项目(\d+)-(\d+).*?项目:([^\n"]+?)\s+批次:([A-Za-z0-9_-]+)', re.S)
    RE_LOAD_CARTRIDGE = re.compile(r'"[^"]*?(\d+)\(\)')
    RE_DILUENT_SCAN = re.compile(r'试剂盘扫码([0-5])')
    RE_BARCODE_CONTENT = re.compile(r'条码内容：([^"]*)')
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
        self.raw_lines: list[str] = []
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
        self.raw_lines = []

        for enc in ('utf-8', 'gbk', 'gb2312', 'gb18030'):
            try:
                with open(path, encoding=enc) as f:
                    lines = f.readlines()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            raise ValueError(f"无法解码文件: {path}")
        self.raw_lines = [line.rstrip("\r\n") for line in lines]

        seen_alarms = set()
        for idx, line in enumerate(lines):
            self._parse_sample(line)
            self._parse_start(line, idx + 1)
            self._parse_finish(line)
            self._parse_motor(line)
            self._parse_alarm(line, seen_alarms, idx + 1)

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
            info.log_range = f"{start_time} - {end_time}"
        elif log_date:
            info.log_time = log_date
            info.log_range = f"{start_time} - {end_time}" if start_time and end_time else ""
        elif start_time and end_time:
            info.log_time = f"{start_time} - {end_time}"
            info.log_range = info.log_time
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
                if mcu_no == "0" and version.startswith("0S") and not info.mcu0_version:
                    info.mcu0_version = version
                elif mcu_no == "1" and version.startswith("1S") and not info.mcu1_version:
                    info.mcu1_version = version
                elif mcu_no == "2" and version.startswith("2S") and not info.mcu2_version:
                    info.mcu2_version = version
                elif mcu_no == "4" and version.startswith("5T") and not info.temp_control_version:
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
        current_self_check = None

        def mark_lines(*line_numbers: int) -> list[int]:
            result = []
            for line_no in line_numbers:
                if line_no and line_no not in result:
                    result.append(line_no)
            return result

        def finalize_self_check():
            nonlocal current_self_check
            if not current_self_check:
                return
            detail = "自检异常" if current_self_check["failed"] else "自检完成"
            self.user_actions.append(UserActionRecord(
                action_time=current_self_check["time"],
                action_type="仪器自检",
                detail=detail,
                source_line=current_self_check["source_line"],
                related_lines=current_self_check["related_lines"],
            ))
            current_self_check = None

        def find_next_barcode(start_idx: int) -> tuple[int, str]:
            for j in range(start_idx + 1, min(len(lines), start_idx + 12)):
                if j != start_idx and self.RE_DILUENT_SCAN.search(lines[j]):
                    break
                if "条码内容" in lines[j]:
                    m = self.RE_BARCODE_CONTENT.search(lines[j])
                    content = (m.group(1) if m else "").strip()
                    content = content.replace("\\r", "").replace("\r", "").strip()
                    return j, content
            return -1, ""

        def find_next_line(start_idx: int, pattern: str, limit: int = 40) -> int:
            for j in range(start_idx + 1, min(len(lines), start_idx + limit)):
                if pattern in lines[j]:
                    return j
            return -1

        for idx, line in enumerate(lines):
            time_str = self._extract_time_prefix(line)
            source_line = idx + 1

            if "system starting" in line:
                user = ""
                related = [source_line]
                for j in range(idx + 1, min(len(lines), idx + 20)):
                    login_match = self.RE_USER_LOGIN.search(lines[j])
                    if login_match:
                        user = login_match.group(1).strip()
                        related.append(j + 1)
                        break
                detail = f"用户登录：{user}" if user else "用户登录：未找到"
                self.user_actions.append(UserActionRecord(
                    action_time=time_str,
                    action_type="仪器开机",
                    detail=detail,
                    source_line=source_line,
                    related_lines=related,
                ))
                continue

            self_check_match = self.RE_SELF_CHECK.search(line)
            if self_check_match:
                mcu_no = self_check_match.group(2)
                if mcu_no == "3":
                    finalize_self_check()
                    current_self_check = {
                        "time": self_check_match.group(1).split(".", 1)[0],
                        "failed": False,
                        "source_line": source_line,
                        "related_lines": [source_line],
                    }
                elif current_self_check:
                    current_self_check["related_lines"].append(source_line)
                    if mcu_no == "1":
                        finalize_self_check()
                continue

            if current_self_check and ("自检错误" in line or self.RE_SELF_CHECK_FAIL.search(line)):
                current_self_check["failed"] = True
                current_self_check["related_lines"].append(source_line)
                continue

            sample_match = self.RE_SAMPLE.search(line)
            if sample_match:
                serial = sample_match.group(2)
                sample = self.samples.get(serial)
                raw_rack = sample_match.group(3)
                rack_display = self._plus_one_rack_pos(raw_rack)
                item_text = "、".join(sample.test_items) if sample and sample.test_items else ""
                detail = f"编号{serial}。项目：{item_text}。样本ID:{sample_match.group(4)}"
                self.user_actions.append(UserActionRecord(
                    action_time=time_str,
                    action_type="急诊位编排" if raw_rack == "6-0" else f"样本架{rack_display}编排",
                    detail=detail,
                    sample_serial=serial,
                    source_line=source_line,
                    related_lines=[source_line],
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
                    source_line=source_line,
                    related_lines=[source_line],
                ))
                pending_project = None
                continue

            diluent_match = self.RE_DILUENT_SCAN.search(line)
            if diluent_match:
                slot = int(diluent_match.group(1)) + 1
                barcode_idx, barcode = find_next_barcode(idx)
                related = mark_lines(source_line, barcode_idx + 1 if barcode_idx >= 0 else 0)
                if barcode and "," in barcode:
                    parts = [part.strip() for part in barcode.split(",")]
                    if len(parts) >= 2 and parts[0] and parts[1]:
                        detail = f"稀释液仓{slot}：稀释液号{parts[0]} 项目{parts[1]}"
                    else:
                        detail = f"稀释液仓{slot}：空"
                else:
                    detail = f"稀释液仓{slot}：空"
                self.user_actions.append(UserActionRecord(
                    action_time=time_str,
                    action_type=f"稀释液仓{slot}装载",
                    detail=detail,
                    source_line=source_line,
                    related_lines=related,
                ))
                continue

            if "耗材盒弹出" in line and "耗材盒弹出完成" not in line:
                finish_idx = find_next_line(idx, "耗材盒弹出完成")
                finish_time = self._extract_time_prefix(lines[finish_idx]) if finish_idx >= 0 else ""
                related = mark_lines(source_line, finish_idx + 1 if finish_idx >= 0 else 0)
                self.user_actions.append(UserActionRecord(
                    action_time=time_str,
                    action_type="耗材更换",
                    detail=f"耗材盒弹出完成时间：{finish_time or '未找到'}",
                    source_line=source_line,
                    related_lines=related,
                ))
                continue

        finalize_self_check()

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
                    finish_time = self._extract_time_prefix(line)
                    s.finish_time = finish_time
                    s.project_abbr = project_abbr
                    s.concentration = concentration
                    s.measure_value = measure_value
                    s.test_results.append({
                        'project_abbr': project_abbr,
                        'concentration': concentration,
                        'measure_value': measure_value,
                        'finish_time': finish_time,
                    })
                pending_serial = None

        # 第二阶段：未完成的样本，检查是否有对应报警
        for serial, sample in self.samples.items():
            if sample.status != "测试完成":
                if any(alarm.sample_num == serial for alarm in self.alarms):
                    sample.status = "异常"

    def _check_action_completeness(self):
        """标准流程功能已移除，不再根据缺失动作改变样本状态。"""
        self.standard_action_sequence = []
        for sample in self.samples.values():
            sample.missing_actions = []

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

    def _parse_start(self, line: str, source_line: int = 0):
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
            source_line=source_line,
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

    def _parse_alarm(self, line: str, seen: set, source_line: int = 0):
        m = self.RE_ALARM.search(line)
        if not m:
            return
        time_str = m.group(1)
        raw_code = m.group(2)
        content = m.group(3)
        reason = m.group(4)
        detail = m.group(5)
        code_match = re.match(
            r'([A-Z]+-[A-Z]+-\d+)(?:-(\d{4,5})([A-Z])?([A-Z]\d{2})?)?',
            raw_code
        )
        if code_match:
            error_code = code_match.group(1)
            sample_num_raw = code_match.group(2)
            action_code = code_match.group(4)
        else:
            error_code = raw_code.split("@", 1)[0]
            sample_num_raw = ""
            action_code = ""

        # 去重：只去除同一秒内完全相同的重复条目（日志可能连续打印两次同一行）
        dedup_key = f"{time_str}-{raw_code}-{content}-{reason}-{detail}"
        if dedup_key in seen:
            return
        seen.add(dedup_key)

        # 取后4位作为样本流水号（无样本号时留空）
        sample_serial = sample_num_raw[-4:] if sample_num_raw else ""
        action_code   = action_code or ""

        alarm = Alarm(
            time_str=time_str,
            error_code=error_code,
            display_code=raw_code,
            sample_num=sample_serial,
            action_code=action_code,
            content=content,
            detail="；".join(part for part in (content, reason, detail) if part),
            source_line=source_line,
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
                font=("Microsoft YaHei", 9), fill=C_TEXT)
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
                                        anchor="s", font=("Consolas", 8),
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
    OPTIONAL_COLS = ["动作时间(ms)", "起始时间", "结束时间", "动作坐标",
                     "理论光栅", "实际光栅", "光栅偏差"]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_CARD)
        self.app = app
        self.col_vars = {}
        self._collapse_var = tk.BooleanVar(value=True)
        self._action_item_map = {}

        # 列显示/隐藏 控制栏
        ctrl_frame = tk.Frame(self, bg=C_BG2)
        ctrl_frame.pack(side=tk.TOP, fill=tk.X, padx=2, pady=2)
        tk.Label(ctrl_frame, text="显示列:", font=("Microsoft YaHei", 9),
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
        self.tree.bind("<Button-3>", self._on_tree_right_click)

        self._rebuild_columns()

    def _block_col_resize(self, event):
        if self.tree.identify_region(event.x, event.y) == "separator":
            return "break"

    def _on_tree_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        act = self._action_item_map.get(item)
        if not act or not act.source_line:
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="查看日志原文",
                         command=lambda: self.app._jump_to_action_raw_line(act))
        menu.tk_popup(event.x_root, event.y_root)

    def _auto_resize_columns(self):
        import tkinter.font as tkfont
        font = tkfont.Font(font=("Microsoft YaHei", 10))
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
        self._repopulate()

    def _repopulate(self):
        self.tree.delete(*self.tree.get_children())
        self._action_item_map = {}
        if not hasattr(self, '_actions'):
            return

        cols = self._get_visible_cols()
        alarm_actions = {a.action_code for a in self._alarms}

        # ── 按时间顺序构建执行批次 ──
        batches = self._build_ordered_batches(self._actions)
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
            # 一级行映射到组内第一个有行号的 Action（右键跳转用）
            first_act = next((a for a in group if a.source_line), None)
            if first_act:
                self._action_item_map[pid] = first_act

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
                self._action_item_map[ciid] = act
                tags = []
                if act.level1 in alarm_actions:
                    tags.append("alarm")
                if abs(act.raster_deviation) >= 3:
                    tags.append("high_dev")
                if tags:
                    self.tree.item(ciid, tags=tuple(tags))

        for level1, group in batches:
            _insert_batch(level1, group)

        _bold = ("Microsoft YaHei", 10, "bold")
        self.tree.tag_configure("alarm", background=C_RED_LIGHT)
        self.tree.tag_configure("high_dev", foreground=C_RED)
        self.tree.tag_configure("summary",
                                background="#c8dcfa", font=_bold)
        # 超时：仅改橙色文字，行背景保持蓝色，与缺失行的整行红色区分
        self.tree.tag_configure("timeout_summary",
                                background="#c8dcfa", foreground="#cc6600", font=_bold)
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
                 bg=C_BG, fg=C_TEXT2, font=("Microsoft YaHei", 9)
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

    def __init__(self, parent, app_dir: str, on_file_loaded=None):
        super().__init__(parent, bg=C_CARD)
        self.app_dir = app_dir
        self.on_file_loaded = on_file_loaded
        self._syncing = False
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
                              font=("Microsoft YaHei", 9))
        self._lbl1.pack(side=tk.LEFT, padx=(0, 14))

        # 参数2
        tk.Button(top, text="载入参数2", command=lambda: self._load_file(1),
                  bg=C_BLUE, fg=C_WHITE, relief="flat", padx=10
                  ).pack(side=tk.LEFT, padx=(0, 4))
        self._lbl2 = tk.Label(top, text="（未载入）", bg=C_BG2, fg=C_TEXT2,
                              font=("Microsoft YaHei", 9))
        self._lbl2.pack(side=tk.LEFT, padx=(0, 14))

        # 仅显示差异 + 差异计数
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        ttk.Checkbutton(top, text="仅显示差异", variable=self.diff_only,
                        command=self._apply_filter).pack(side=tk.LEFT)
        self._diff_lbl = tk.Label(top, text="差异数：—", bg=C_BG2, fg=C_TEXT2,
                                  font=("Microsoft YaHei", 9))
        self._diff_lbl.pack(side=tk.LEFT, padx=8)

        # 搜索栏
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(top, text="搜索:", bg=C_BG2, fg=C_TEXT2,
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        search_entry = ttk.Entry(top, textvariable=self._search_var, width=18)
        search_entry.pack(side=tk.LEFT, padx=(4, 2))
        search_entry.bind("<Return>", lambda e: self._do_search())
        self._search_var.trace("w", lambda *a: self._do_search())
        tk.Button(top, text="✕", command=self._clear_search,
                  bg=C_BG2, fg=C_TEXT2, relief="flat", font=("Microsoft YaHei", 9),
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
                                font=("Microsoft YaHei", 10, "bold"))
        self.tree.tag_configure("group_diff", background="#ffe8e8",
                                font=("Microsoft YaHei", 10, "bold"))

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
        if self.on_file_loaded and not self._syncing:
            self.on_file_loaded(self, idx, self.data[idx], name)

    def set_shared_data(self, data, file_names):
        """同步另一参数视图已载入的数据。"""
        self._syncing = True
        try:
            self.data = data
            self.file_names = file_names
            self._lbl1.config(text=file_names[0] or "（未载入）")
            self._lbl2.config(text=file_names[1] or "（未载入）")
            self._load_action_names()
            self._load_param_notes()
            self._populate()
        finally:
            self._syncing = False

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

    def _make_param_item(self, label: str, v1, v2) -> dict:
        return {
            "label": label,
            "v1": v1,
            "v2": v2,
            "s1": self._fmt(v1),
            "s2": self._fmt(v2),
            "diff": (v1 is None or v2 is None or self._is_diff(v1, v2)),
            "note": self._param_notes.get(label, ""),
        }

    def _value_param_items(self, v1, v2, preferred_labels=None, prefix: str = "") -> list[dict]:
        if isinstance(v1 or v2, dict):
            sub1, sub2 = (v1 or {}), (v2 or {})
            items = []
            for k in self._merged_keys(sub1, sub2):
                label = f"{prefix}{k}" if prefix else k
                items.append(self._make_param_item(label, sub1.get(k), sub2.get(k)))
            return items

        if isinstance(v1 or v2, list):
            arr1 = v1 if isinstance(v1, list) else []
            arr2 = v2 if isinstance(v2, list) else []
            n = max(len(arr1), len(arr2), len(preferred_labels or []))
            items = []
            for i in range(n):
                base_label = preferred_labels[i] if preferred_labels and i < len(preferred_labels) else f"参数{i + 1}"
                label = f"{prefix}{base_label}" if prefix else base_label
                e1 = arr1[i] if i < len(arr1) else None
                e2 = arr2[i] if i < len(arr2) else None
                items.append(self._make_param_item(label, e1, e2))
            return items

        return [self._make_param_item(prefix or "值", v1, v2)]

    def _build_param_sections(self) -> list[dict]:
        """生成整机参数的统一展示模型，供树形视图和表格视图复用。"""
        d1 = self.data[0] or {}
        d2 = self.data[1] or {}
        all_keys = list(dict.fromkeys(list(d1.keys()) + list(d2.keys())))
        motor_keys = [
            k for k in all_keys
            if isinstance(d1.get(k) or d2.get(k), dict)
            and "动作参数" in (d1.get(k) or d2.get(k) or {})
            and k not in ("ADP",)
        ]
        detection_mainboard_keys = ("动态算法", "放大倍数", "曲线翻转")
        sections: list[dict] = []

        system_items = [
            self._make_param_item(k, d1.get(k), d2.get(k))
            for k in self.SYSTEM_KEYS
            if k in all_keys
        ]
        sections.append({"title": "系统信息", "grouped": False, "items": system_items})

        adp1 = (d1.get("ADP") or {}).get("动作参数", [])
        adp2 = (d2.get("ADP") or {}).get("动作参数", [])
        adp_items = []
        for i in range(max(len(adp1), len(adp2))):
            key = f"ADP_{i}"
            named = key in self._action_names
            v1 = adp1[i] if i < len(adp1) else None
            v2 = adp2[i] if i < len(adp2) else None
            if not named and (v1 or 0) == 0 and (v2 or 0) == 0:
                continue
            adp_items.append(self._make_param_item(self._action_names.get(key, f"参数[{i}]"), v1, v2))
        sections.append({"title": "加样器参数", "grouped": False, "items": adp_items})

        motor_groups = []
        for mkey in motor_keys:
            m1, m2 = (d1.get(mkey) or {}), (d2.get(mkey) or {})
            items = [
                self._make_param_item(k, m1.get(k), m2.get(k))
                for k in self._merged_keys(m1, m2)
                if k != "动作参数"
            ]
            motor_groups.append({"title": mkey, "items": items})
        sections.append({"title": "电机参数", "grouped": True, "groups": motor_groups})

        coord_groups = []
        coord1, coord2 = (d1.get("坐标") or {}), (d2.get("坐标") or {})
        if coord1 or coord2:
            coord_items = []
            for pt in self._merged_keys(coord1, coord2):
                xyz1 = coord1.get(pt, [None, None, None])
                xyz2 = coord2.get(pt, [None, None, None])
                for i, axis in enumerate(["X", "Y", "Z"]):
                    v1 = xyz1[i] if xyz1 and i < len(xyz1) else None
                    v2 = xyz2[i] if xyz2 and i < len(xyz2) else None
                    coord_items.append(self._make_param_item(f"{pt}  {axis}", v1, v2))
            coord_groups.append({"title": "坐标定位点", "items": coord_items})
        for mkey in motor_keys:
            m1, m2 = (d1.get(mkey) or {}), (d2.get(mkey) or {})
            arr1, arr2 = m1.get("动作参数", []), m2.get("动作参数", [])
            items = []
            for i in range(max(len(arr1), len(arr2), 20)):
                key = f"{mkey}_{i}"
                named = key in self._action_names
                v1 = arr1[i] if i < len(arr1) else None
                v2 = arr2[i] if i < len(arr2) else None
                if not named and (v1 or 0) == 0 and (v2 or 0) == 0:
                    continue
                items.append(self._make_param_item(self._action_names.get(key, f"参数[{i}]"), v1, v2))
            coord_groups.append({"title": f"{mkey}  动作参数", "items": items})
        sections.append({"title": "坐标参数", "grouped": True, "groups": coord_groups})

        temp_groups = []
        for k in [k for k in all_keys if "温控" in k]:
            tc1, tc2 = (d1.get(k) or {}), (d2.get(k) or {})
            temp_groups.append({
                "title": k,
                "items": [self._make_param_item(fk, tc1.get(fk), tc2.get(fk)) for fk in self._merged_keys(tc1, tc2)]
            })
        sections.append({"title": "温控参数", "grouped": True, "groups": temp_groups})

        sample_groups = []
        sample_key = next((k for k in ("样本容器", "管径参数") if k in all_keys), None)
        if sample_key:
            sc1, sc2 = (d1.get(sample_key) or {}), (d2.get(sample_key) or {})
            for i in range(1, 10):
                tube = f"tube{i}"
                if tube in sc1 or tube in sc2:
                    sample_groups.append({
                        "title": tube,
                        "items": self._value_param_items(sc1.get(tube), sc2.get(tube),
                                                         ["参数1", "参数2", "参数3", "参数4"])
                    })
        sections.append({"title": "样本容器", "grouped": True, "groups": sample_groups})

        detection_items = []
        channel_key = next((k for k in ("检测参数", "通道参数") if k in all_keys), None)
        ch1, ch2 = (d1.get(channel_key) or {}) if channel_key else {}, (d2.get(channel_key) or {}) if channel_key else {}
        for sk in self._merged_keys(ch1, ch2):
            sv1, sv2 = ch1.get(sk), ch2.get(sk)
            if isinstance(sv1 or sv2, (dict, list)):
                detection_items.extend(self._value_param_items(sv1, sv2, prefix=f"{sk} "))
            else:
                detection_items.append(self._make_param_item(sk, sv1, sv2))
        mb1, mb2 = d1.get("主板参数", {}), d2.get("主板参数", {})
        for k in detection_mainboard_keys:
            if k in mb1 or k in mb2:
                detection_items.append(self._make_param_item(k, mb1.get(k), mb2.get(k)))
        sections.append({"title": "检测参数", "grouped": False, "items": detection_items})

        return sections

    def _insert_group(self, parent, label: str, open_: bool = False) -> str:
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

    def _merged_keys(self, d1: dict, d2: dict):
        return dict.fromkeys(list(d1.keys()) + list(d2.keys()))

    def _insert_dict_group(self, parent: str, label: str, d1: dict, d2: dict) -> bool:
        g = self._insert_group(parent, label)
        g_diff = False
        for k in self._merged_keys(d1, d2):
            g_diff |= self._insert_leaf(g, k, d1.get(k), d2.get(k))
        self._mark_group_diff(g, g_diff)
        return g_diff

    def _insert_value_children(self, parent: str, v1, v2, preferred_labels=None) -> bool:
        has_diff = False
        if isinstance(v1 or v2, dict):
            sub1, sub2 = (v1 or {}), (v2 or {})
            for k in self._merged_keys(sub1, sub2):
                has_diff |= self._insert_leaf(parent, k, sub1.get(k), sub2.get(k))
            return has_diff

        if isinstance(v1 or v2, list):
            arr1 = v1 if isinstance(v1, list) else []
            arr2 = v2 if isinstance(v2, list) else []
            n = max(len(arr1), len(arr2), len(preferred_labels or []))
            for i in range(n):
                label = preferred_labels[i] if preferred_labels and i < len(preferred_labels) else f"参数{i + 1}"
                e1 = arr1[i] if i < len(arr1) else None
                e2 = arr2[i] if i < len(arr2) else None
                has_diff |= self._insert_leaf(parent, label, e1, e2)
            return has_diff

        return self._insert_leaf(parent, "值", v1, v2)

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

        detection_mainboard_keys = ("动态算法", "放大倍数", "曲线翻转")

        # ── 1. 系统信息 ──────────────────────────────────────────────────
        g = self._insert_group("", "系统信息")
        g_diff = False
        for k in self.SYSTEM_KEYS:
            if k in all_keys:
                g_diff |= self._insert_leaf(g, k, d1.get(k), d2.get(k))
        self._mark_group_diff(g, g_diff)
        self._diff_count += sum(1 for _, d in self._leaf_iids[-len(self.SYSTEM_KEYS):] if d)

        # ── 2. 加样器参数 ────────────────────────────────────────────────
        adp1 = (d1.get("ADP") or {}).get("动作参数", [])
        adp2 = (d2.get("ADP") or {}).get("动作参数", [])
        if adp1 or adp2:
            g = self._insert_group("", "加样器参数")
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

        # ── 3. 电机参数组（除动作参数）────────────────────────────────────
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

        # ── 4. 坐标参数（坐标定位点 + 各电机动作参数）────────────────────
        g_coord = self._insert_group("", "坐标参数")
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

        # ── 5. 温控参数 ──────────────────────────────────────────────────
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

        # ── 6. 样本容器（tube1-tube9）───────────────────────────────────
        sample_key = next((k for k in ("样本容器", "管径参数") if k in all_keys), None)
        if sample_key:
            sc1, sc2 = (d1.get(sample_key) or {}), (d2.get(sample_key) or {})
            g_sample = self._insert_group("", "样本容器")
            g_sample_diff = False
            for i in range(1, 10):
                tube = f"tube{i}"
                if tube not in sc1 and tube not in sc2:
                    continue
                tg = self._insert_group(g_sample, tube)
                tg_diff = self._insert_value_children(tg, sc1.get(tube), sc2.get(tube),
                                                      ["参数1", "参数2", "参数3", "参数4"])
                self._mark_group_diff(tg, tg_diff)
                g_sample_diff |= tg_diff
            self._mark_group_diff(g_sample, g_sample_diff)

        # ── 7. 检测参数（原通道参数 + 主板检测项）──────────────────────────
        channel_key = next((k for k in ("检测参数", "通道参数") if k in all_keys), None)
        ch1, ch2 = (d1.get(channel_key) or {}) if channel_key else {}, (d2.get(channel_key) or {}) if channel_key else {}
        mb1, mb2 = d1.get("主板参数", {}), d2.get("主板参数", {})
        has_detection_from_mb = any(k in mb1 or k in mb2 for k in detection_mainboard_keys)
        if ch1 or ch2 or has_detection_from_mb:
            g_detection = self._insert_group("", "检测参数")
            detection_diff = False
            if ch1 or ch2:
                for sk in self._merged_keys(ch1, ch2):
                    sv1, sv2 = ch1.get(sk), ch2.get(sk)
                    if isinstance(sv1 or sv2, (dict, list)):
                        sg = self._insert_group(g_detection, sk)
                        sg_diff = self._insert_value_children(sg, sv1, sv2)
                        self._mark_group_diff(sg, sg_diff)
                        detection_diff |= sg_diff
                    else:
                        detection_diff |= self._insert_leaf(g_detection, sk, sv1, sv2)
            for k in detection_mainboard_keys:
                if k in mb1 or k in mb2:
                    detection_diff |= self._insert_leaf(g_detection, k, mb1.get(k), mb2.get(k))
            self._mark_group_diff(g_detection, detection_diff)

        # ── 其他顶层 key：跳过日志和已归类项，避免生成多余一级列表 ───────
        handled = (
            set(self.SYSTEM_KEYS) | {"主板参数", "ADP", "坐标", "日志"} |
            {"样本容器", "管径参数", "检测参数", "通道参数"} |
            set(motor_keys) | set(tc_keys)
        )
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


class ParamsTableView(ParamsView):
    """整机参数表格展开视图。"""

    ROW_H = 28
    TITLE_H = 32
    GAP_H = 12
    FIRST_COL_W = 65
    MIN_TABLE_W = 420
    MIN_COL_W = 68

    def _build_ui(self):
        top = tk.Frame(self, bg=C_BG2)
        top.pack(fill=tk.X, padx=6, pady=4)

        tk.Button(top, text="导入整机参数", command=self._load_machine_params,
                  bg=C_BLUE, fg=C_WHITE, relief="flat", padx=10
                  ).pack(side=tk.LEFT, padx=(0, 4))
        self._lbl1 = tk.Label(top, text="（未载入）", bg=C_BG2, fg=C_TEXT2,
                              font=("Microsoft YaHei", 9))
        self._lbl1.pack(side=tk.LEFT, padx=(0, 14))

        self._diff_lbl = tk.Label(top, text="参数数：—", bg=C_BG2, fg=C_TEXT2,
                                  font=("Microsoft YaHei", 9))
        self._diff_lbl.pack(side=tk.LEFT, padx=8)

        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(top, text="搜索:", bg=C_BG2, fg=C_TEXT2,
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        search_entry = ttk.Entry(top, textvariable=self._search_var, width=18)
        search_entry.pack(side=tk.LEFT, padx=(4, 2))
        search_entry.bind("<Return>", lambda e: self._do_search())
        self._search_var.trace("w", lambda *a: self._do_search())
        tk.Button(top, text="✕", command=self._clear_search,
                  bg=C_BG2, fg=C_TEXT2, relief="flat", font=("Microsoft YaHei", 9),
                  cursor="hand2", padx=2).pack(side=tk.LEFT)

        body = tk.Frame(self, bg=C_CARD)
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        self.canvas = tk.Canvas(body, bg=C_WHITE, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.canvas.yview)
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)
        self.canvas.bind("<Configure>", lambda e: self._populate())
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<Double-1>", self._on_canvas_double_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self._drawn_cells = []
        self._section_cells = []
        self._expanded_sections = set()
        self._search_match_key = None
        self._font_normal = tkfont.Font(family="Microsoft YaHei", size=9)
        self._font_bold = tkfont.Font(family="Microsoft YaHei", size=9, weight="bold")
        self._font_title = tkfont.Font(family="Microsoft YaHei", size=10, weight="bold")

    def _load_machine_params(self):
        path = filedialog.askopenfilename(
            title="导入整机参数",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        import json as _json
        with open(path, encoding="utf-8") as f:
            self.data[0] = _json.load(f)
        self.data[1] = None
        name = os.path.basename(path)
        self.file_names = [name, ""]
        self._lbl1.config(text=name)
        self._expanded_sections.clear()
        self._populate()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _text(self, x, y, text, width, font=None, fill=C_TEXT, anchor="w"):
        text = "" if text is None else str(text)
        return self.canvas.create_text(
            x, y, text=text, anchor=anchor, fill=fill,
            font=font or ("Microsoft YaHei", 9)
        )

    def _draw_cell(self, x, y, w, h, text="", fill=C_WHITE, outline=C_BORDER,
                   font=None, text_fill=C_TEXT, cell_meta=None):
        self.canvas.create_rectangle(x, y, x + w, y + h, fill=fill, outline=outline)
        self._text(x + 8, y + h / 2, text, w - 12, font=font, fill=text_fill)
        if cell_meta:
            cell_meta["bbox"] = (x, y, x + w, y + h)
            self._drawn_cells.append(cell_meta)

    def _item_key(self, item):
        return (item["label"], item["s1"], item["s2"])

    def _measure_text(self, text, bold=False):
        font = self._font_bold if bold else self._font_normal
        return font.measure("" if text is None else str(text))

    def _column_widths(self, items):
        widths = []
        for item in items:
            content_w = max(
                self._measure_text(item["label"], bold=True),
                self._measure_text(item["s1"]),
            ) + 18
            widths.append(max(self.MIN_COL_W, content_w))
        return widths

    def _table_width(self, items, table_w):
        return max(table_w, self.FIRST_COL_W + sum(self._column_widths(items)))

    def _draw_section_header(self, y, title, width, expanded):
        fill = "#dfe8f8" if expanded else "#eef2fa"
        self.canvas.create_rectangle(0, y, width, y + self.TITLE_H, fill=fill, outline=C_BORDER)
        marker = "▼" if expanded else "▶"
        self._text(10, y + self.TITLE_H / 2, f"{marker} {title}", width - 20,
                   font=("Microsoft YaHei", 10, "bold"))
        self._section_cells.append({"title": title, "bbox": (0, y, width, y + self.TITLE_H)})
        return y + self.TITLE_H

    def _draw_table_rows(self, y, title, items, table_w, query, draw_title=True):
        if query:
            q = query.lower()
            items = [
                p for p in items
                if q in p["label"].lower()
                or q in p["s1"].lower()
                or q in title.lower()
            ]
        if not items:
            if draw_title:
                return self._draw_section_header(y, title, table_w, True)
            return y

        col_widths = self._column_widths(items)
        width = max(table_w, self.FIRST_COL_W + sum(col_widths))
        header_fill = "#eef2fa"
        search_fill = "#fff0a6"

        if draw_title:
            y = self._draw_section_header(y, title, width, True)

        row_labels = [title, "参数值"]
        rows = [
            [p["label"] for p in items],
            [p["s1"] for p in items],
        ]
        for row_idx in range(2):
            first_fill = header_fill if row_idx == 0 else "#f7f9fd"
            self._draw_cell(0, y, self.FIRST_COL_W, self.ROW_H, row_labels[row_idx],
                            fill=first_fill, font=("Microsoft YaHei", 9, "bold"))
            x = self.FIRST_COL_W
            for col_idx, item in enumerate(items):
                col_w = col_widths[col_idx]
                fill = C_WHITE
                if row_idx == 0:
                    fill = header_fill
                if self._search_match_key == self._item_key(item) and row_idx == 0:
                    fill = search_fill
                self._draw_cell(
                    x, y, col_w, self.ROW_H, rows[row_idx][col_idx], fill=fill,
                    font=("Microsoft YaHei", 9, "bold") if row_idx == 0 else None,
                    cell_meta={
                        "item": item,
                        "text": " ".join([title, item["label"], item["s1"], item.get("note", "")]),
                    }
                )
                x += col_w
            y += self.ROW_H
        return y

    def _draw_grouped_section(self, y, section, table_w, query):
        groups = section.get("groups", [])
        if query:
            q = query.lower()
            filtered = []
            for g in groups:
                items = [
                    p for p in g["items"]
                    if q in p["label"].lower()
                    or q in p["s1"].lower()
                    or q in g["title"].lower()
                    or q in section["title"].lower()
                ]
                if items:
                    filtered.append({"title": g["title"], "items": items})
            groups = filtered

        if section["title"] == "电机参数":
            width = self._motor_table_width(groups, table_w)
        else:
            width = max(table_w, max([self._table_width(g["items"], table_w) for g in groups] or [table_w]))
        is_expanded = bool(query) or section["title"] in self._expanded_sections
        y = self._draw_section_header(y, section["title"], width, is_expanded)
        if not is_expanded:
            return y

        if section["title"] == "电机参数":
            return self._draw_motor_table(y, groups, width)

        for group_index, group in enumerate(groups):
            self.canvas.create_rectangle(0, y, width, y + self.ROW_H, fill="#edf4ff", outline=C_BORDER)
            self._text(14, y + self.ROW_H / 2, group["title"], width - 20,
                       font=("Microsoft YaHei", 9, "bold"))
            y += self.ROW_H
            y = self._draw_table_rows(y, group["title"], group["items"], width, "", draw_title=False)
        return y

    def _motor_headers(self, groups):
        headers = []
        seen = set()
        for group in groups:
            for item in group.get("items", []):
                label = item["label"]
                if label not in seen:
                    headers.append(label)
                    seen.add(label)
        return headers

    def _motor_column_widths(self, groups, headers):
        widths = []
        for label in headers:
            values = [label]
            for group in groups:
                item = next((p for p in group.get("items", []) if p["label"] == label), None)
                if item:
                    values.append(item["s1"])
            content_w = max(self._measure_text(v, bold=(v == label)) for v in values) + 18
            widths.append(max(self.MIN_COL_W, content_w))
        return widths

    def _motor_first_col_width(self, groups):
        labels = ["电机名称"] + [g["title"] for g in groups]
        return max(self.FIRST_COL_W, max(self._measure_text(label, bold=True) for label in labels) + 18)

    def _motor_table_width(self, groups, table_w):
        headers = self._motor_headers(groups)
        return max(table_w, self._motor_first_col_width(groups) + sum(self._motor_column_widths(groups, headers)))

    def _draw_motor_table(self, y, groups, table_w):
        if not groups:
            return y
        headers = self._motor_headers(groups)
        col_widths = self._motor_column_widths(groups, headers)
        first_col_w = self._motor_first_col_width(groups)
        header_fill = "#eef2fa"
        search_fill = "#fff0a6"

        self._draw_cell(0, y, first_col_w, self.ROW_H, "电机名称",
                        fill=header_fill, font=("Microsoft YaHei", 9, "bold"))
        x = first_col_w
        for i, label in enumerate(headers):
            col_w = col_widths[i]
            self._draw_cell(x, y, col_w, self.ROW_H, label,
                            fill=header_fill, font=("Microsoft YaHei", 9, "bold"))
            x += col_w
        y += self.ROW_H

        for group in groups:
            item_by_label = {item["label"]: item for item in group.get("items", [])}
            self._draw_cell(0, y, first_col_w, self.ROW_H, group["title"],
                            fill="#f7f9fd", font=("Microsoft YaHei", 9, "bold"))
            x = first_col_w
            for i, label in enumerate(headers):
                col_w = col_widths[i]
                item = item_by_label.get(label, {"label": label, "s1": "", "s2": "", "note": ""})
                fill = search_fill if self._search_match_key == self._item_key(item) else C_WHITE
                self._draw_cell(
                    x, y, col_w, self.ROW_H, item["s1"], fill=fill,
                    cell_meta={
                        "item": item,
                        "text": " ".join([group["title"], item["label"], item["s1"], item.get("note", "")]),
                    }
                )
                x += col_w
            y += self.ROW_H
        return y

    def _populate(self):
        if not hasattr(self, "canvas"):
            return
        self.canvas.delete("all")
        self._drawn_cells = []
        self._section_cells = []
        sections = self._build_param_sections()
        param_count = sum(
            1 for section in sections
            for item in (section.get("items", []) if not section.get("grouped") else
                         [p for g in section.get("groups", []) for p in g.get("items", [])])
        )
        self._diff_lbl.config(text=f"参数数：{param_count}")

        query = self._search_var.get().strip().lower()
        y = 0
        table_w = max(self.MIN_TABLE_W, self.canvas.winfo_width() - 2)
        for section in sections:
            before = y
            if section.get("grouped"):
                y = self._draw_grouped_section(y, section, table_w, query)
            else:
                is_expanded = bool(query) or section["title"] in self._expanded_sections
                items = section.get("items", [])
                width = self._table_width(items, table_w)
                if is_expanded:
                    y = self._draw_table_rows(y, section["title"], items, table_w, query)
                else:
                    y = self._draw_section_header(y, section["title"], width, False)
            if y != before:
                y += self.GAP_H
        self.canvas.configure(scrollregion=(0, 0, max(table_w, self.canvas.bbox("all")[2] if self.canvas.bbox("all") else table_w), y))
        if self._search_match_key:
            self._scroll_to_match()

    def _apply_filter(self):
        self._populate()

    def _do_search(self):
        query = self._search_var.get().strip().lower()
        self._search_match_key = None
        if query:
            for section in self._build_param_sections():
                if section.get("grouped"):
                    pools = [p for g in section.get("groups", []) for p in g.get("items", [])]
                else:
                    pools = section.get("items", [])
                for item in pools:
                    text = " ".join([item["label"], item["s1"], item.get("note", "")]).lower()
                    if query in text:
                        self._search_match_key = self._item_key(item)
                        break
                if self._search_match_key:
                    break
        self._populate()

    def _scroll_to_match(self):
        for cell in self._drawn_cells:
            if self._item_key(cell.get("item")) == self._search_match_key:
                x1, y1, x2, y2 = cell["bbox"]
                bbox = self.canvas.bbox("all")
                if bbox:
                    total_h = max(1, bbox[3] - bbox[1])
                    total_w = max(1, bbox[2] - bbox[0])
                    self.canvas.yview_moveto(max(0, (y1 - 40) / total_h))
                    self.canvas.xview_moveto(max(0, (x1 - self.FIRST_COL_W) / total_w))
                break

    def _on_canvas_click(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        for cell in self._section_cells:
            x1, y1, x2, y2 = cell["bbox"]
            if x1 <= x <= x2 and y1 <= y <= y2:
                title = cell["title"]
                if title in self._expanded_sections:
                    self._expanded_sections.remove(title)
                else:
                    self._expanded_sections.add(title)
                self._populate()
                return

    def _on_canvas_double_click(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        for cell in self._drawn_cells:
            x1, y1, x2, y2 = cell["bbox"]
            if x1 <= x <= x2 and y1 <= y <= y2:
                item = cell.get("item")
                note = item.get("note", "") if item else ""
                if note:
                    messagebox.showinfo("备注说明", f"{item['label']}\n\n{note}", parent=self)
                return


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
            views = self.param_view if isinstance(self.param_view, (list, tuple)) else [self.param_view]
            for view in views:
                view._load_param_notes()
                view._populate()

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self, text="双击「备注说明」列可编辑，备注将同步显示在整机参数对比列表中",
                 bg=C_BG, fg=C_TEXT2, font=("Microsoft YaHei", 9)
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
                 fg=C_TEXT2, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=8)

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
        self._syncing_sample_selection = False
        self._raw_font_size = 11

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
                    font=("Microsoft YaHei", 9))
        s.map("TCheckbutton", background=[("active", C_BG2)])
        s.configure("Treeview",
                    rowheight=24, font=("Microsoft YaHei", 10),
                    background=C_WHITE, fieldbackground=C_WHITE,
                    foreground=C_TEXT, bordercolor=C_BORDER)
        s.configure("Treeview.Heading",
                    font=("Microsoft YaHei", 10, "bold"),
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
                                  bg=C_BLUE, fg=C_WHITE, font=("Microsoft YaHei", 10),
                                  relief="flat", padx=12, pady=3, cursor="hand2")
        self.load_btn.pack(side=tk.LEFT, padx=8)
        tk.Button(toolbar, text="动作备注", command=self._open_theory_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 10),
                  relief="flat", padx=12, pady=3, bd=1, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(toolbar, text="电机备注", command=self._open_action_name_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 10),
                  relief="flat", padx=12, pady=3, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(toolbar, text="参数备注", command=self._open_param_notes_dialog,
                  bg=C_WHITE, fg=C_TEXT, font=("Microsoft YaHei", 10),
                  relief="flat", padx=12, pady=3, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=4)
        # 文件名显示
        self.file_label = tk.Label(toolbar, text="未载入文件", bg=C_BG2,
                                   fg=C_TEXT2, font=("Microsoft YaHei", 9))
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
        self.params_notebook = ttk.Notebook(self.params_tab)
        self.params_notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        self.params_table_page = tk.Frame(self.params_notebook, bg=C_CARD)
        self.params_compare_page = tk.Frame(self.params_notebook, bg=C_CARD)
        self.params_notebook.add(self.params_table_page, text="整机参数")
        self.params_notebook.add(self.params_compare_page, text="参数对比")

        self._param_table_view = ParamsTableView(self.params_table_page, app_dir)
        self._param_table_view.pack(fill=tk.BOTH, expand=True)
        self._param_view = ParamsView(self.params_compare_page, app_dir)
        self._param_view.pack(fill=tk.BOTH, expand=True)
        self._param_views = [self._param_view, self._param_table_view]

        self._build_instrument_tab()
        self._build_sample_tab()
        self._build_alarm_tab()

    def _sync_param_views(self, source, idx, data, name):
        for view in getattr(self, "_param_views", []):
            if view is source:
                continue
            view.set_shared_data(source.data, source.file_names)

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
                                      bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 11, "bold"),
                                      padx=8, pady=6)
        summary_frame.pack(fill=tk.X, padx=10, pady=(6, 4))
        self.instrument_value_labels = {}
        field_columns = [
            [
                ("日志时间", "log_time"),
                ("仪器序列号", "device_serial"),
                ("用户程序版本", "user_program_version"),
                ("中位机版本", "mid_version"),
            ],
            [
                ("MCU3版本", "mcu2_version"),
                ("MCU2版本", "mcu1_version"),
                ("MCU1版本", "mcu0_version"),
                ("温控版本", "temp_control_version"),
            ],
            [
                ("申请样本次数", "request_count"),
                ("累计检测次数", "detect_count"),
                ("累计开盖次数", "open_cap_count"),
                ("日志编排数", "current_arrangement_count"),
            ],
            [
                ("4G CCID", "ccid"),
                ("信号强度", "signal_strength"),
            ],
        ]
        for col, fields in enumerate(field_columns):
            base_col = col * 2
            for row, (label_text, key) in enumerate(fields):
                tk.Label(summary_frame, text=f"{label_text}:", bg=C_CARD, fg=C_TEXT,
                         font=("Microsoft YaHei", 10, "bold")).grid(
                             row=row, column=base_col, sticky="w", padx=(0, 6), pady=2)
                value_label = tk.Label(summary_frame, text="-", bg=C_CARD, fg=C_TEXT2,
                                       font=("Microsoft YaHei", 10), anchor="w")
                value_label.grid(row=row, column=base_col + 1, sticky="w", padx=(0, 14), pady=2)
                self.instrument_value_labels[key] = value_label
        for col in range(8):
            summary_frame.grid_columnconfigure(col, weight=1 if col % 2 else 0)

        action_frame = tk.LabelFrame(self.instrument_tab, text="用户动作",
                                     bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 11, "bold"),
                                     padx=6, pady=6)
        action_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))
        action_split = tk.Frame(action_frame, bg=C_CARD)
        action_split.pack(fill=tk.BOTH, expand=True)
        action_split.grid_columnconfigure(0, weight=1, uniform="instrument_halves")
        action_split.grid_columnconfigure(1, weight=1, uniform="instrument_halves")
        action_split.grid_rowconfigure(0, weight=1)

        action_left = tk.Frame(action_split, bg=C_CARD)
        action_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        filter_bar = tk.Frame(action_left, bg=C_CARD)
        filter_bar.pack(fill=tk.X, pady=(0, 4))
        tk.Label(filter_bar, text="筛选：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.user_action_filter_vars = {}
        self.user_action_all_var = tk.BooleanVar(value=True)
        tk.Checkbutton(filter_bar, text="全选", variable=self.user_action_all_var,
                       command=self._toggle_all_user_action_filters,
                       bg=C_CARD, fg=C_TEXT, selectcolor=C_CARD,
                       activebackground=C_CARD,
                       font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=(0, 4))
        for text in ("开机", "自检", "编排", "装载弹夹", "稀释液装载", "耗材更换"):
            var = tk.BooleanVar(value=True)
            self.user_action_filter_vars[text] = var
            tk.Checkbutton(filter_bar, text=text, variable=var,
                           command=self._on_user_action_filter_changed,
                           bg=C_CARD, fg=C_TEXT, selectcolor=C_CARD,
                           activebackground=C_CARD,
                           font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=2)

        tree_frame, self.user_action_tree = self._create_tree(
            action_left,
            ("动作时间", "动作", "详情"),
            widths=[82, 140, 380],
            height=8,
        )
        self.user_action_tree.column("动作时间", width=82, minwidth=76, anchor="center", stretch=False)
        self.user_action_tree.column("动作", width=140, minwidth=120, anchor="center", stretch=False)
        self.user_action_tree.column("详情", width=380, minwidth=300, anchor="w", stretch=False)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.user_action_tree.bind("<<TreeviewSelect>>", self._on_user_action_select)
        self.user_action_tree.bind("<Configure>", self._resize_user_action_columns, add="+")
        self.user_action_tree.tag_configure("action_error", foreground=C_RED)

        preview_frame = tk.Frame(action_split, bg=C_CARD)
        preview_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        preview_frame.grid_rowconfigure(1, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        search_bar = tk.Frame(preview_frame, bg=C_CARD)
        search_bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Label(search_bar, text="原文搜索：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.raw_search_var1 = tk.StringVar()
        self.raw_search_var2 = tk.StringVar()
        self.raw_search_mode_var = tk.StringVar(value="或")
        self.raw_search_entry1 = ttk.Entry(search_bar, textvariable=self.raw_search_var1, width=16)
        self.raw_search_entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 2))
        self.raw_search_mode = ttk.Combobox(
            search_bar,
            textvariable=self.raw_search_mode_var,
            values=("或", "和"),
            state="readonly",
            width=4,
        )
        self.raw_search_mode.pack(side=tk.LEFT, padx=2)
        self.raw_search_entry2 = ttk.Entry(search_bar, textvariable=self.raw_search_var2, width=16)
        self.raw_search_entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.raw_search_entry1.bind("<Return>", lambda e: self._search_raw_text("down"))
        self.raw_search_entry2.bind("<Return>", lambda e: self._search_raw_text("down"))
        tk.Button(search_bar, text="向上", command=lambda: self._search_raw_text("up"),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=8).pack(side=tk.LEFT, padx=(4, 0))
        tk.Button(search_bar, text="向下", command=lambda: self._search_raw_text("down"),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=8).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Separator(search_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        tk.Button(search_bar, text="A+", command=lambda: self._adjust_raw_font_size(1),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=6).pack(side=tk.LEFT)
        tk.Button(search_bar, text="A-", command=lambda: self._adjust_raw_font_size(-1),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=6).pack(side=tk.LEFT, padx=(2, 0))

        raw_text_frame = tk.Frame(preview_frame, bg=C_CARD)
        raw_text_frame.grid(row=1, column=0, sticky="nsew")
        raw_text_frame.grid_rowconfigure(0, weight=1)
        raw_text_frame.grid_columnconfigure(0, weight=1)
        self.raw_text = tk.Text(
            raw_text_frame,
            wrap="none",
            bg=C_WHITE,
            fg=C_TEXT,
            font=("Consolas", 11),
            relief="solid",
            bd=1,
        )
        raw_vsb = ttk.Scrollbar(raw_text_frame, orient=tk.VERTICAL, command=self.raw_text.yview)
        raw_hsb = ttk.Scrollbar(raw_text_frame, orient=tk.HORIZONTAL, command=self.raw_text.xview)
        self.raw_text.configure(yscrollcommand=raw_vsb.set, xscrollcommand=raw_hsb.set)
        self.raw_text.grid(row=0, column=0, sticky="nsew")
        raw_vsb.grid(row=0, column=1, sticky="ns")
        raw_hsb.grid(row=1, column=0, sticky="ew")
        self.raw_text.tag_configure("source_highlight", background="#fff59d")
        self.raw_text.tag_configure("search_highlight", background="#c8e6c9")
        self.raw_text.config(state="disabled")

    def _build_sample_tab(self):
        sample_list_frame = tk.LabelFrame(self.sample_tab, text="样本编排信息",
                                          bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 11, "bold"),
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
                 bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(stat_frame, textvariable=self._stat_done_var,
                 bg=C_CARD, fg="#1a7f3c", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(stat_frame, textvariable=self._stat_error_var,
                 bg=C_CARD, fg=C_RED, font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(stat_frame, textvariable=self._stat_unknown_var,
                 bg=C_CARD, fg=C_TEXT2, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(stat_frame, text="显示：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
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
                           font=("Microsoft YaHei", 10),
                           activebackground=C_CARD).pack(side=tk.LEFT, padx=4)

        # ── R-01 样本搜索框 ──
        search_frame = tk.Frame(sample_list_frame, bg=C_CARD)
        search_frame.pack(fill=tk.X, pady=(0, 4))
        tk.Label(search_frame, text="搜索：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self._sample_search_var = tk.StringVar()
        self._sample_search_var.trace_add("write", lambda *_: self._apply_sample_filter())
        ttk.Entry(search_frame, textvariable=self._sample_search_var,
                  width=24).pack(side=tk.LEFT, padx=4)

        # ── 左右分栏：左侧样本列表 / 右侧项目测试结果 ──
        split_frame = tk.Frame(sample_list_frame, bg=C_CARD)
        split_frame.pack(fill=tk.BOTH, expand=True)
        split_frame.grid_columnconfigure(0, weight=3)
        split_frame.grid_columnconfigure(1, weight=1)
        split_frame.grid_rowconfigure(0, weight=1)

        left_frame = tk.Frame(split_frame, bg=C_CARD)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        tree_frame, self.sample_tree = self._create_tree(
            left_frame,
            ("编排时间", "编号", "样本ID", "样本类型", "样本位置", "测试数", "开盖", "摇匀", "项目", "样本状态"),
            widths=[80, 60, 130, 60, 60, 45, 45, 45, 100, 80],
            height=7,
        )
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.sample_tree.bind("<<TreeviewSelect>>", self._on_sample_tree_select)

        right_panel = tk.LabelFrame(split_frame, text="项目测试结果",
                                    bg=C_CARD, fg=C_TEXT,
                                    font=("Microsoft YaHei", 10, "bold"),
                                    padx=4, pady=4)
        right_panel.grid(row=0, column=1, sticky="nsew")

        result_tree_frame, self.result_tree = self._create_tree(
            right_panel,
            ("项目缩写", "项目模式", "完成时间", "浓度", "测试值"),
            widths=[80, 80, 80, 70, 90],
            height=7,
        )
        result_tree_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Separator(self.sample_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        detail_frame = tk.LabelFrame(self.sample_tab, text="样本动作详情",
                                     bg=C_CARD, fg=C_TEXT, font=("Microsoft YaHei", 11, "bold"),
                                     padx=8, pady=8)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        tab_frame = tk.Frame(detail_frame, bg=C_BG2)
        tab_frame.pack(fill=tk.X, pady=(0, 6))

        self._view_var = tk.StringVar(value="table")
        tk.Radiobutton(tab_frame, text="表格视图", variable=self._view_var,
                       value="table", command=self._switch_view,
                       bg=C_BG2, fg=C_TEXT, font=("Microsoft YaHei", 10),
                       selectcolor=C_BG2, indicatoron=0, padx=15, pady=3,
                       relief="flat").pack(side=tk.LEFT, padx=2, pady=2)
        tk.Radiobutton(tab_frame, text="日志原文", variable=self._view_var,
                       value="rawlog", command=self._switch_view,
                       bg=C_BG2, fg=C_TEXT, font=("Microsoft YaHei", 10),
                       selectcolor=C_BG2, indicatoron=0, padx=15, pady=3,
                       relief="flat").pack(side=tk.LEFT, padx=2, pady=2)

        self.view_container = tk.Frame(detail_frame, bg=C_CARD)
        self.view_container.pack(fill=tk.BOTH, expand=True)
        self.timeline = TimelineCanvas(self.view_container, self)
        self.table_view = TableView(self.view_container, self)
        self.table_view.pack(fill=tk.BOTH, expand=True)

        # 日志原文视图
        self.sample_raw_frame = tk.Frame(self.view_container, bg=C_CARD)
        self.sample_raw_frame.grid_rowconfigure(1, weight=1)
        self.sample_raw_frame.grid_columnconfigure(0, weight=1)

        sample_raw_search_bar = tk.Frame(self.sample_raw_frame, bg=C_CARD)
        sample_raw_search_bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Label(sample_raw_search_bar, text="原文搜索：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.sample_raw_search_var1 = tk.StringVar()
        self.sample_raw_search_var2 = tk.StringVar()
        self.sample_raw_search_mode_var = tk.StringVar(value="或")
        self.sample_raw_search_entry1 = ttk.Entry(sample_raw_search_bar,
                                                   textvariable=self.sample_raw_search_var1, width=16)
        self.sample_raw_search_entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 2))
        ttk.Combobox(sample_raw_search_bar, textvariable=self.sample_raw_search_mode_var,
                     values=("或", "和"), state="readonly", width=4).pack(side=tk.LEFT, padx=2)
        self.sample_raw_search_entry2 = ttk.Entry(sample_raw_search_bar,
                                                   textvariable=self.sample_raw_search_var2, width=16)
        self.sample_raw_search_entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.sample_raw_search_entry1.bind("<Return>", lambda e: self._search_sample_raw_text("down"))
        self.sample_raw_search_entry2.bind("<Return>", lambda e: self._search_sample_raw_text("down"))
        tk.Button(sample_raw_search_bar, text="向上",
                  command=lambda: self._search_sample_raw_text("up"),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=8).pack(side=tk.LEFT, padx=(4, 0))
        tk.Button(sample_raw_search_bar, text="向下",
                  command=lambda: self._search_sample_raw_text("down"),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=8).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Separator(sample_raw_search_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        tk.Button(sample_raw_search_bar, text="A+", command=lambda: self._adjust_raw_font_size(1),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=6).pack(side=tk.LEFT)
        tk.Button(sample_raw_search_bar, text="A-", command=lambda: self._adjust_raw_font_size(-1),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=6).pack(side=tk.LEFT, padx=(2, 0))

        raw_text_frame = tk.Frame(self.sample_raw_frame, bg=C_CARD)
        raw_text_frame.grid(row=1, column=0, sticky="nsew")
        raw_text_frame.grid_rowconfigure(0, weight=1)
        raw_text_frame.grid_columnconfigure(0, weight=1)
        self.sample_raw_text = tk.Text(
            raw_text_frame, wrap="none", bg=C_WHITE, fg=C_TEXT,
            font=("Consolas", 11), relief="solid", bd=1,
        )
        srvsb = ttk.Scrollbar(raw_text_frame, orient=tk.VERTICAL, command=self.sample_raw_text.yview)
        srhsb = ttk.Scrollbar(raw_text_frame, orient=tk.HORIZONTAL, command=self.sample_raw_text.xview)
        self.sample_raw_text.configure(yscrollcommand=srvsb.set, xscrollcommand=srhsb.set)
        self.sample_raw_text.grid(row=0, column=0, sticky="nsew")
        srvsb.grid(row=0, column=1, sticky="ns")
        srhsb.grid(row=1, column=0, sticky="ew")
        self.sample_raw_text.tag_configure("source_highlight", background="#fff59d")
        self.sample_raw_text.tag_configure("search_highlight", background="#c8e6c9")
        self.sample_raw_text.config(state="disabled")

    def _build_alarm_tab(self):
        alarm_split = tk.Frame(self.alarm_tab, bg=C_CARD)
        alarm_split.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        alarm_split.grid_columnconfigure(0, weight=1, uniform="alarm_halves")
        alarm_split.grid_columnconfigure(1, weight=1, uniform="alarm_halves")
        alarm_split.grid_rowconfigure(0, weight=1)

        alarm_left = tk.Frame(alarm_split, bg=C_CARD)
        alarm_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        tk.Label(alarm_left, text="异常信息", bg=C_CARD, fg=C_RED,
                 font=("Microsoft YaHei", 11, "bold")).pack(anchor="w", pady=(0, 4))
        filter_frame = tk.Frame(alarm_left, bg=C_CARD)
        filter_frame.pack(fill=tk.X, pady=(0, 4))
        tk.Label(filter_frame, text="显示：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.alarm_show_info_var = tk.BooleanVar(value=True)
        self.alarm_show_error_var = tk.BooleanVar(value=True)
        for text, var, fg in [
            ("信息(C)", self.alarm_show_info_var, C_TEXT),
            ("异常(G)", self.alarm_show_error_var, C_RED),
        ]:
            tk.Checkbutton(filter_frame, text=text, variable=var,
                           command=self._populate_alarm_tree,
                           bg=C_CARD, fg=fg, selectcolor=C_CARD,
                           activebackground=C_CARD,
                           font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=4)
        alarm_tree_frame, self.alarm_tree = self._create_tree(
            alarm_left,
            ("时间", "异常编号", "详情"),
            widths=[78, 220, 300],
        )
        self.alarm_tree.column("时间", width=78, minwidth=72, anchor="center", stretch=False)
        self.alarm_tree.column("异常编号", width=220, minwidth=180, anchor="center", stretch=False)
        self.alarm_tree.column("详情", width=300, minwidth=220, anchor="w", stretch=True)
        alarm_tree_frame.pack(fill=tk.BOTH, expand=True)
        self.alarm_tree.bind("<<TreeviewSelect>>", self._on_alarm_select)

        alarm_right = tk.Frame(alarm_split, bg=C_CARD)
        alarm_right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        alarm_right.grid_rowconfigure(1, weight=1)
        alarm_right.grid_columnconfigure(0, weight=1)
        search_bar = tk.Frame(alarm_right, bg=C_CARD)
        search_bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Label(search_bar, text="原文搜索：", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.alarm_raw_search_var1 = tk.StringVar()
        self.alarm_raw_search_var2 = tk.StringVar()
        self.alarm_raw_search_mode_var = tk.StringVar(value="或")
        self.alarm_raw_search_entry1 = ttk.Entry(search_bar, textvariable=self.alarm_raw_search_var1, width=16)
        self.alarm_raw_search_entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 2))
        ttk.Combobox(search_bar, textvariable=self.alarm_raw_search_mode_var,
                     values=("或", "和"), state="readonly", width=4).pack(side=tk.LEFT, padx=2)
        self.alarm_raw_search_entry2 = ttk.Entry(search_bar, textvariable=self.alarm_raw_search_var2, width=16)
        self.alarm_raw_search_entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.alarm_raw_search_entry1.bind("<Return>", lambda e: self._search_alarm_raw_text("down"))
        self.alarm_raw_search_entry2.bind("<Return>", lambda e: self._search_alarm_raw_text("down"))
        tk.Button(search_bar, text="向上", command=lambda: self._search_alarm_raw_text("up"),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=8).pack(side=tk.LEFT, padx=(4, 0))
        tk.Button(search_bar, text="向下", command=lambda: self._search_alarm_raw_text("down"),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=8).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Separator(search_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        tk.Button(search_bar, text="A+", command=lambda: self._adjust_raw_font_size(1),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=6).pack(side=tk.LEFT)
        tk.Button(search_bar, text="A-", command=lambda: self._adjust_raw_font_size(-1),
                  bg=C_WHITE, fg=C_TEXT, relief="flat", padx=6).pack(side=tk.LEFT, padx=(2, 0))

        raw_text_frame = tk.Frame(alarm_right, bg=C_CARD)
        raw_text_frame.grid(row=1, column=0, sticky="nsew")
        raw_text_frame.grid_rowconfigure(0, weight=1)
        raw_text_frame.grid_columnconfigure(0, weight=1)
        self.alarm_raw_text = tk.Text(
            raw_text_frame, wrap="none", bg=C_WHITE, fg=C_TEXT,
            font=("Consolas", 11), relief="solid", bd=1,
        )
        raw_vsb = ttk.Scrollbar(raw_text_frame, orient=tk.VERTICAL, command=self.alarm_raw_text.yview)
        raw_hsb = ttk.Scrollbar(raw_text_frame, orient=tk.HORIZONTAL, command=self.alarm_raw_text.xview)
        self.alarm_raw_text.configure(yscrollcommand=raw_vsb.set, xscrollcommand=raw_hsb.set)
        self.alarm_raw_text.grid(row=0, column=0, sticky="nsew")
        raw_vsb.grid(row=0, column=1, sticky="ns")
        raw_hsb.grid(row=1, column=0, sticky="ew")
        self.alarm_raw_text.tag_configure("source_highlight", background="#fff59d")
        self.alarm_raw_text.tag_configure("search_highlight", background="#c8e6c9")
        self.alarm_raw_text.config(state="disabled")

    def _switch_view(self):
        view = self._view_var.get()
        if view == "rawlog":
            self.table_view.pack_forget()
            self.sample_raw_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self.sample_raw_frame.pack_forget()
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
        self._load_alarm_raw_text()
        self._load_sample_raw_text()

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
            self._populate_result_tree("")
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
        self.table_view.set_data(actions, sample_alarms, self.parser.theory_times)
        self._populate_result_tree(serial)

    def _on_alarm_select(self, event):
        sel = self.alarm_tree.selection()
        if not sel:
            return
        item_id = sel[0]
        alarm = getattr(self, "_alarm_item_map", {}).get(item_id)
        if not alarm:
            return

        serials = list(self.parser.samples.keys())
        if alarm.sample_num in serials:
            self._select_sample(alarm.sample_num)
            t_ms = time_to_ms(alarm.time_str + ".000") if '.' not in alarm.time_str else time_to_ms(alarm.time_str)
            self.timeline.scroll_to_time(t_ms)
            self.table_view.highlight_action(alarm.action_code)
        if alarm.source_line:
            self._highlight_alarm_raw_lines([alarm.source_line])

    def highlight_alarm(self, alarm: Alarm):
        """从时间轴点击报警后，高亮异常信息列表"""
        for item in self.alarm_tree.get_children():
            row_alarm = getattr(self, "_alarm_item_map", {}).get(item)
            if row_alarm and row_alarm.error_code == alarm.error_code and row_alarm.sample_num == alarm.sample_num:
                self.alarm_tree.selection_set(item)
                self.alarm_tree.see(item)
                if row_alarm.source_line:
                    self._highlight_alarm_raw_lines([row_alarm.source_line])
                break

    def _populate_instrument_tab(self):
        info = self.parser.instrument_info
        for key, label in self.instrument_value_labels.items():
            value = getattr(info, key, "")
            label.config(text=str(value) if value not in ("", None) else "未找到")
        self._load_raw_text()
        self._populate_user_action_tree()

    def _action_category(self, action_type: str) -> str:
        if "开机" in action_type:
            return "开机"
        if "自检" in action_type:
            return "自检"
        if "编排" in action_type:
            return "编排"
        if "装载弹夹" in action_type:
            return "装载弹夹"
        if "稀释液" in action_type:
            return "稀释液装载"
        if "耗材更换" in action_type:
            return "耗材更换"
        return "全部"

    def _toggle_all_user_action_filters(self):
        selected = self.user_action_all_var.get()
        for var in self.user_action_filter_vars.values():
            var.set(selected)
        self._populate_user_action_tree()

    def _on_user_action_filter_changed(self):
        if hasattr(self, "user_action_all_var"):
            self.user_action_all_var.set(all(var.get() for var in self.user_action_filter_vars.values()))
        self._populate_user_action_tree()

    def _populate_user_action_tree(self):
        self.user_action_tree.delete(*self.user_action_tree.get_children())
        self._user_action_item_map = {}
        for idx, record in enumerate(self.parser.user_actions):
            category = self._action_category(record.action_type)
            filter_vars = getattr(self, "user_action_filter_vars", {})
            if category != "全部" and filter_vars and not filter_vars.get(category, tk.BooleanVar(value=True)).get():
                continue
            tags = ("action_error",) if record.detail == "自检异常" else ()
            item_id = f"useraction-{idx}"
            self.user_action_tree.insert("", tk.END, iid=item_id, values=(
                record.action_time,
                record.action_type,
                record.detail,
            ), tags=tags)
            self._user_action_item_map[item_id] = record
        self.user_action_tree.after_idle(self._resize_user_action_columns)

    def _resize_user_action_columns(self, event=None):
        if not hasattr(self, "user_action_tree"):
            return
        tree_width = event.width if event is not None else self.user_action_tree.winfo_width()
        if tree_width <= 1:
            return

        scrollbar_width = 22
        available_width = max(602, tree_width - scrollbar_width)
        min_time_width = 82
        min_action_width = 140
        min_detail_width = 380
        min_total = min_time_width + min_action_width + min_detail_width
        extra_width = max(0, available_width - min_total)

        time_width = min_time_width + int(extra_width * 0.16)
        action_width = min_action_width + int(extra_width * 0.24)
        detail_width = available_width - time_width - action_width

        self.user_action_tree.column("动作时间", width=time_width)
        self.user_action_tree.column("动作", width=action_width)
        self.user_action_tree.column("详情", width=detail_width)

    def _load_raw_text(self):
        if not hasattr(self, "raw_text"):
            return
        self.raw_text.config(state="normal")
        self.raw_text.delete("1.0", tk.END)
        if self.parser.raw_lines:
            self.raw_text.insert("1.0", "\n".join(self.parser.raw_lines))
        self.raw_text.tag_remove("source_highlight", "1.0", tk.END)
        self.raw_text.tag_remove("search_highlight", "1.0", tk.END)
        self.raw_text.config(state="disabled")

    def _load_alarm_raw_text(self):
        if not hasattr(self, "alarm_raw_text"):
            return
        self.alarm_raw_text.config(state="normal")
        self.alarm_raw_text.delete("1.0", tk.END)
        if self.parser.raw_lines:
            self.alarm_raw_text.insert("1.0", "\n".join(self.parser.raw_lines))
        self.alarm_raw_text.tag_remove("source_highlight", "1.0", tk.END)
        self.alarm_raw_text.tag_remove("search_highlight", "1.0", tk.END)
        self.alarm_raw_text.config(state="disabled")

    def _load_sample_raw_text(self):
        if not hasattr(self, "sample_raw_text"):
            return
        self.sample_raw_text.config(state="normal")
        self.sample_raw_text.delete("1.0", tk.END)
        if self.parser.raw_lines:
            self.sample_raw_text.insert("1.0", "\n".join(self.parser.raw_lines))
        self.sample_raw_text.tag_remove("source_highlight", "1.0", tk.END)
        self.sample_raw_text.tag_remove("search_highlight", "1.0", tk.END)
        self.sample_raw_text.config(state="disabled")

    def _highlight_sample_raw_lines(self, line_numbers: list[int]):
        if not hasattr(self, "sample_raw_text"):
            return
        self._highlight_lines_in_text(self.sample_raw_text, line_numbers)

    def _search_sample_raw_text(self, direction: str = "down"):
        if not hasattr(self, "sample_raw_text"):
            return
        keyword1 = self.sample_raw_search_var1.get().strip()
        keyword2 = self.sample_raw_search_var2.get().strip()
        mode = self.sample_raw_search_mode_var.get()
        if not keyword1 and not keyword2:
            messagebox.showinfo("搜索提示", "请输入搜索内容")
            return
        match_idx = self._find_raw_match_line(keyword1, keyword2, mode, direction,
                                               self.sample_raw_text)
        if match_idx < 0:
            messagebox.showinfo("搜索提示", "未搜索到相关内容")
            return
        line_no = match_idx + 1
        self.sample_raw_text.config(state="normal")
        self._highlight_search_keywords(line_no, [keyword1, keyword2], self.sample_raw_text)
        self.sample_raw_text.see(f"{line_no}.0")
        self.sample_raw_text.mark_set(tk.INSERT, f"{line_no}.0")
        self.sample_raw_text.config(state="disabled")

    def _jump_to_action_raw_line(self, act):
        self._view_var.set("rawlog")
        self._switch_view()
        self._highlight_sample_raw_lines([act.source_line])

    def _adjust_raw_font_size(self, delta: int):
        self._raw_font_size = max(8, min(28, self._raw_font_size + delta))
        new_font = ("Consolas", self._raw_font_size)
        for attr in ("raw_text", "alarm_raw_text", "sample_raw_text"):
            w = getattr(self, attr, None)
            if w:
                w.config(font=new_font)

    def _on_user_action_select(self, event):
        sel = self.user_action_tree.selection()
        if not sel:
            return
        record = getattr(self, "_user_action_item_map", {}).get(sel[0])
        if record:
            self._highlight_raw_lines(record.related_lines or [record.source_line])

    def _highlight_raw_lines(self, line_numbers: list[int]):
        if not hasattr(self, "raw_text"):
            return
        self._highlight_lines_in_text(self.raw_text, line_numbers)

    def _highlight_alarm_raw_lines(self, line_numbers: list[int]):
        if not hasattr(self, "alarm_raw_text"):
            return
        self._highlight_lines_in_text(self.alarm_raw_text, line_numbers)

    def _highlight_lines_in_text(self, text_widget: tk.Text, line_numbers: list[int]):
        text_widget.config(state="normal")
        text_widget.tag_remove("source_highlight", "1.0", tk.END)
        valid_lines = [line_no for line_no in line_numbers if line_no > 0]
        for line_no in valid_lines:
            text_widget.tag_add("source_highlight", f"{line_no}.0", f"{line_no}.end")
        if valid_lines:
            text_widget.see(f"{valid_lines[0]}.0")
            text_widget.mark_set(tk.INSERT, f"{valid_lines[0]}.0")
        text_widget.config(state="disabled")

    def _line_has_keyword(self, line: str, keyword: str) -> bool:
        return keyword.lower() in line.lower()

    def _raw_line_matches_search(self, idx: int, keyword1: str, keyword2: str, mode: str) -> bool:
        lines = self.parser.raw_lines
        line = lines[idx]
        has1 = bool(keyword1) and self._line_has_keyword(line, keyword1)
        has2 = bool(keyword2) and self._line_has_keyword(line, keyword2)
        if mode == "或":
            return has1 or has2
        if not keyword1 or not keyword2:
            return has1 or has2
        if not has1 and not has2:
            return False
        start = max(0, idx - 10)
        end = min(len(lines), idx + 11)
        if has1:
            return any(self._line_has_keyword(lines[j], keyword2) for j in range(start, end))
        return any(self._line_has_keyword(lines[j], keyword1) for j in range(start, end))

    def _find_raw_match_line(self, keyword1: str, keyword2: str, mode: str,
                             direction: str, text_widget: tk.Text = None) -> int:
        total = len(self.parser.raw_lines)
        if total == 0:
            return -1
        text_widget = text_widget or self.raw_text
        try:
            current_line = int(float(text_widget.index(tk.INSERT).split(".", 1)[0])) - 1
        except (ValueError, tk.TclError):
            current_line = 0

        if direction == "up":
            order = list(range(current_line - 1, -1, -1)) + list(range(total - 1, current_line, -1))
        else:
            order = list(range(current_line + 1, total)) + list(range(0, current_line + 1))
        for idx in order:
            if self._raw_line_matches_search(idx, keyword1, keyword2, mode):
                return idx
        return -1

    def _highlight_search_keywords(self, line_no: int, keywords: list[str], text_widget: tk.Text = None):
        text_widget = text_widget or self.raw_text
        text_widget.tag_remove("search_highlight", "1.0", tk.END)
        for keyword in keywords:
            if not keyword:
                continue
            start = f"{line_no}.0"
            while True:
                pos = text_widget.search(keyword, start, stopindex=f"{line_no}.end", nocase=True)
                if not pos:
                    break
                end = f"{pos}+{len(keyword)}c"
                text_widget.tag_add("search_highlight", pos, end)
                start = end

    def _search_raw_text(self, direction: str = "down"):
        if not hasattr(self, "raw_text"):
            return
        keyword1 = self.raw_search_var1.get().strip()
        keyword2 = self.raw_search_var2.get().strip()
        mode = self.raw_search_mode_var.get()
        if not keyword1 and not keyword2:
            messagebox.showinfo("搜索提示", "请输入搜索内容")
            return
        match_idx = self._find_raw_match_line(keyword1, keyword2, mode, direction, self.raw_text)
        if match_idx < 0:
            messagebox.showinfo("搜索提示", "未搜索到相关内容")
            return

        line_no = match_idx + 1
        self.raw_text.config(state="normal")
        self._highlight_search_keywords(line_no, [keyword1, keyword2], self.raw_text)
        self.raw_text.see(f"{line_no}.0")
        self.raw_text.mark_set(tk.INSERT, f"{line_no}.0")
        self.raw_text.config(state="disabled")

    def _search_alarm_raw_text(self, direction: str = "down"):
        if not hasattr(self, "alarm_raw_text"):
            return
        keyword1 = self.alarm_raw_search_var1.get().strip()
        keyword2 = self.alarm_raw_search_var2.get().strip()
        mode = self.alarm_raw_search_mode_var.get()
        if not keyword1 and not keyword2:
            messagebox.showinfo("搜索提示", "请输入搜索内容")
            return
        match_idx = self._find_raw_match_line(keyword1, keyword2, mode, direction, self.alarm_raw_text)
        if match_idx < 0:
            messagebox.showinfo("搜索提示", "未搜索到相关内容")
            return

        line_no = match_idx + 1
        self.alarm_raw_text.config(state="normal")
        self._highlight_search_keywords(line_no, [keyword1, keyword2], self.alarm_raw_text)
        self.alarm_raw_text.see(f"{line_no}.0")
        self.alarm_raw_text.mark_set(tk.INSERT, f"{line_no}.0")
        self.alarm_raw_text.config(state="disabled")

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
        search_text  = getattr(self, "_sample_search_var", tk.StringVar()).get().strip().lower()

        self.sample_tree.delete(*self.sample_tree.get_children())
        for serial in sorted(self.parser.samples.keys()):
            sample = self.parser.samples[serial]

            # 状态筛选
            if sample.status == "测试完成" and not show_done:
                continue
            if sample.status == "异常" and not show_error:
                continue
            if sample.status not in ("测试完成", "异常") and not show_unknown:
                continue

            item_text = "、".join(sample.test_items[:2])

            # 搜索过滤（流水号、样本ID、项目）
            if search_text and not any(
                search_text in (val or "").lower()
                for val in [sample.serial, sample.sample_id, item_text]
            ):
                continue
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
                sample.status,
            ))

    def _populate_result_tree(self, serial: str):
        if not hasattr(self, "result_tree"):
            return
        self.result_tree.delete(*self.result_tree.get_children())
        if not serial:
            return
        sample = self.parser.samples.get(serial)
        if not sample:
            return
        results = getattr(sample, 'test_results', None) or []
        if not results and sample.project_abbr:
            results = [{
                'project_abbr': sample.project_abbr,
                'concentration': sample.concentration,
                'measure_value': sample.measure_value,
                'finish_time': sample.finish_time,
            }]
        for r in results:
            self.result_tree.insert("", tk.END, values=(
                r.get('project_abbr', ''),
                sample.mode,
                r.get('finish_time', ''),
                r.get('concentration', ''),
                r.get('measure_value', ''),
            ))

    def _apply_sample_filter(self):
        """筛选复选框变化时刷新样本列表（保留当前选中样本）。"""
        sel = self.sample_tree.selection()
        prev_serial = sel[0] if sel else ""
        self._populate_sample_tree()
        if prev_serial and prev_serial in self.sample_tree.get_children():
            self.sample_tree.selection_set(prev_serial)
            self.sample_tree.see(prev_serial)
            self._populate_result_tree(prev_serial)

    def _populate_alarm_tree(self):
        self.alarm_tree.delete(*self.alarm_tree.get_children())
        self._alarm_item_map = {}
        show_info = self.alarm_show_info_var.get() if hasattr(self, "alarm_show_info_var") else True
        show_error = self.alarm_show_error_var.get() if hasattr(self, "alarm_show_error_var") else True
        for idx, alarm in enumerate(self.parser.alarms):
            kind = alarm.error_code[:1]
            if kind == "C" and not show_info:
                continue
            if kind == "G" and not show_error:
                continue
            item_id = f"alarm-{idx}"
            description = alarm.detail or alarm.content or alarm.error_code
            self.alarm_tree.insert("", tk.END, iid=item_id, values=(
                alarm.time_str,
                alarm.display_code,
                description,
            ))
            self._alarm_item_map[item_id] = alarm

    def _open_params_tab(self):
        self.main_notebook.select(self.params_tab)

    def _open_param_notes_dialog(self):
        app_dir = _app_dir()
        ParamNotesDialog(self.root, app_dir, param_view=getattr(self, "_param_views", [self._param_view]))

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
