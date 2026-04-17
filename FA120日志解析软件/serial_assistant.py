"""
和迈串口调试助手 - Serial Port Debugging Assistant
Requirements: pip install pyserial
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import threading
import time
from datetime import datetime

# ── Color palette (Google Material) ─────────────────────────────────────────
C_BG       = "#f1f3f4"   # light blue-grey page background
C_BG2      = "#e8eaed"   # slightly darker for bars / card headers
C_CARD     = "#ffffff"   # card / text area background
C_BORDER   = "#dadce0"
C_BLUE     = "#1a73e8"
C_BLUE_HV  = "#1557b0"
C_GREEN    = "#34a853"
C_RED      = "#ea4335"
C_TEXT     = "#202124"
C_TEXT2    = "#5f6368"
C_WHITE    = "#ffffff"


class SerialAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("和迈串口调试助手 v1.0")
        self.root.geometry("980x680")
        self.root.minsize(820, 560)
        self.root.configure(bg=C_BG)   # light blue-grey

        self.serial_port = None
        self.is_connected = False
        self.receive_thread = None
        self.rx_count = 0
        self.tx_count = 0
        self.running = False
        self._rx_buf = bytearray()
        self._auto_send_job = None

        self._apply_style()
        self._build_ui()
        self._refresh_ports()

    # ── Global ttk style ─────────────────────────────────────────────────────

    def _apply_style(self):
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure("TCombobox",
                    fieldbackground=C_WHITE,
                    background=C_WHITE,
                    foreground=C_TEXT,
                    bordercolor=C_BORDER,
                    arrowcolor=C_TEXT2,
                    relief="flat")
        s.map("TCombobox",
              fieldbackground=[("readonly", C_WHITE)],
              bordercolor=[("focus", C_BLUE)])
        s.configure("TSeparator", background=C_BORDER)
        s.configure("TCheckbutton",
                    background=C_BG,
                    foreground=C_TEXT,
                    font=("Microsoft YaHei", 8))
        s.map("TCheckbutton", background=[("active", C_BG)])

    # ── UI Construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._divider(self.root)
        self._build_settings_bar()
        self._divider(self.root)
        self._build_body()
        self._build_statusbar()

    # ── Header ───────────────────────────────────────────────────────────────

    def _build_header(self):
        header = tk.Frame(self.root, bg=C_BG2, height=52)
        header.pack(fill=tk.X, padx=16)
        header.pack_propagate(False)

        # Brand
        tk.Label(
            header,
            text="和迈串口调试助手",
            font=("Microsoft YaHei", 15, "bold"),
            fg=C_BLUE, bg=C_BG2
        ).pack(side=tk.LEFT, pady=12)

        # Connection status badge (right)
        self.status_badge = tk.Label(
            header,
            text="● 未连接",
            font=("Microsoft YaHei", 9),
            fg=C_TEXT2, bg=C_BG2
        )
        self.status_badge.pack(side=tk.RIGHT, pady=12, padx=4)

    # ── Settings bar (horizontal) ─────────────────────────────────────────────

    def _build_settings_bar(self):
        bar = tk.Frame(self.root, bg=C_BG2, height=64)
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)

        inner = tk.Frame(bar, bg=C_BG2)
        inner.pack(side=tk.LEFT, padx=12, pady=10)

        def combo_unit(parent, label, var, values, width=9):
            f = tk.Frame(parent, bg=C_BG2)
            f.pack(side=tk.LEFT, padx=(0, 10))
            tk.Label(f, text=label, font=("Microsoft YaHei", 8),
                     fg=C_TEXT2, bg=C_BG2).pack(anchor=tk.W)
            cb = ttk.Combobox(f, textvariable=var, values=values,
                               width=width, state="readonly")
            cb.pack()
            return cb

        # Port
        port_f = tk.Frame(inner, bg=C_BG2)
        port_f.pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(port_f, text="串口", font=("Microsoft YaHei", 8),
                 fg=C_TEXT2, bg=C_BG2).pack(anchor=tk.W)
        port_inner = tk.Frame(port_f, bg=C_BG2)
        port_inner.pack()
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(port_inner, textvariable=self.port_var,
                                        width=8, state="readonly")
        self.port_combo.pack(side=tk.LEFT)
        tk.Button(port_inner, text="↺", command=self._refresh_ports,
                  font=("Arial", 9), relief=tk.FLAT, bg=C_BG2, fg=C_TEXT2,
                  activebackground=C_BORDER, cursor="hand2", bd=0,
                  padx=2).pack(side=tk.LEFT, padx=(2, 0))

        # Baud
        self.baud_var = tk.StringVar(value="115200")
        combo_unit(inner, "波特率", self.baud_var,
                   ["1200","2400","4800","9600","19200","38400",
                    "57600","115200","230400","460800","921600"], width=9)

        # Data bits
        self.databits_var = tk.StringVar(value="8")
        combo_unit(inner, "数据位", self.databits_var, ["5","6","7","8"], width=4)

        # Stop bits
        self.stopbits_var = tk.StringVar(value="1")
        combo_unit(inner, "停止位", self.stopbits_var, ["1","1.5","2"], width=4)

        # Parity
        self.parity_var = tk.StringVar(value="None")
        combo_unit(inner, "校验位", self.parity_var,
                   ["None","Even","Odd","Mark","Space"], width=6)

        # Encoding
        self.encoding_var = tk.StringVar(value="GBK")
        combo_unit(inner, "编码", self.encoding_var,
                   ["GBK","UTF-8","ASCII","Latin-1"], width=7)

        # Connect button (right side)
        btn_frame = tk.Frame(bar, bg=C_BG2)
        btn_frame.pack(side=tk.RIGHT, padx=16, pady=12)
        self.connect_btn = tk.Button(
            btn_frame, text="打开串口",
            command=self._toggle_connection,
            font=("Microsoft YaHei", 9, "bold"),
            bg=C_BLUE, fg=C_WHITE,
            activebackground=C_BLUE_HV, activeforeground=C_WHITE,
            relief=tk.FLAT, cursor="hand2",
            padx=18, pady=6, bd=0
        )
        self.connect_btn.pack()

    # ── Body (receive + send) ─────────────────────────────────────────────────

    def _build_body(self):
        body = tk.Frame(self.root, bg=C_BG)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        self._build_rx_card(body)
        self._build_send_card(body)

    def _build_rx_card(self, parent):
        card = tk.Frame(parent, bg=C_CARD, bd=1,
                        highlightbackground=C_BORDER,
                        highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # Card header row
        hdr = tk.Frame(card, bg=C_BG2, height=34)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)

        tk.Label(hdr, text="接收区", font=("Microsoft YaHei", 9, "bold"),
                 fg=C_TEXT, bg=C_BG2).pack(side=tk.LEFT, padx=10, pady=6)

        # Checkboxes (right of header)
        opts = tk.Frame(hdr, bg=C_BG2)
        opts.pack(side=tk.LEFT, padx=6, pady=4)

        self.rx_hex_var = tk.BooleanVar(value=False)
        self.rx_timestamp_var = tk.BooleanVar(value=False)
        self.rx_newline_var = tk.BooleanVar(value=True)
        self.auto_scroll_var = tk.BooleanVar(value=True)

        for text, var in [("HEX显示", self.rx_hex_var),
                           ("时间戳", self.rx_timestamp_var),
                           ("自动换行", self.rx_newline_var),
                           ("自动滚动", self.auto_scroll_var)]:
            ttk.Checkbutton(opts, text=text, variable=var).pack(side=tk.LEFT, padx=4)

        # Action buttons (far right)
        btns = tk.Frame(hdr, bg=C_BG2)
        btns.pack(side=tk.RIGHT, padx=8, pady=4)

        self._outline_btn(btns, "保存", self._save_rx, C_BLUE).pack(side=tk.RIGHT, padx=(4, 0))
        self._outline_btn(btns, "清空", self._clear_rx, C_RED).pack(side=tk.RIGHT)

        # Text area
        self.rx_text = scrolledtext.ScrolledText(
            card,
            font=("Consolas", 10),
            bg=C_CARD, fg=C_TEXT,
            insertbackground=C_TEXT,
            wrap=tk.WORD,
            state=tk.DISABLED,
            relief=tk.FLAT,
            borderwidth=0,
            padx=8, pady=6
        )
        self.rx_text.pack(fill=tk.BOTH, expand=True)

        self.rx_text.tag_configure("rx",  foreground=C_TEXT)
        self.rx_text.tag_configure("tx",  foreground=C_BLUE)
        self.rx_text.tag_configure("sys", foreground=C_TEXT2)
        self.rx_text.tag_configure("ts",  foreground=C_GREEN)

    def _build_send_card(self, parent):
        card = tk.Frame(parent, bg=C_CARD, bd=1,
                        highlightbackground=C_BORDER,
                        highlightthickness=1)
        card.pack(fill=tk.X)

        # Card header row
        hdr = tk.Frame(card, bg=C_BG2, height=34)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)

        tk.Label(hdr, text="发送区", font=("Microsoft YaHei", 9, "bold"),
                 fg=C_TEXT, bg=C_BG2).pack(side=tk.LEFT, padx=10, pady=6)

        # Send options
        opts = tk.Frame(hdr, bg=C_BG2)
        opts.pack(side=tk.LEFT, padx=6, pady=4)

        self.tx_hex_var = tk.BooleanVar(value=False)
        self.tx_newline_var = tk.BooleanVar(value=True)
        self.auto_send_var = tk.BooleanVar(value=False)
        self.auto_interval_var = tk.StringVar(value="1000")

        ttk.Checkbutton(opts, text="HEX发送", variable=self.tx_hex_var).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(opts, text="自动换行(\\r\\n)", variable=self.tx_newline_var).pack(side=tk.LEFT, padx=4)

        # Auto send
        auto_f = tk.Frame(opts, bg=C_BG2)
        auto_f.pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(auto_f, text="定时(ms):", variable=self.auto_send_var,
                        command=self._toggle_auto_send).pack(side=tk.LEFT)
        tk.Entry(auto_f, textvariable=self.auto_interval_var,
                 width=5, font=("Arial", 8),
                 bg=C_WHITE, fg=C_TEXT,
                 relief=tk.FLAT, bd=1,
                 highlightbackground=C_BORDER,
                 highlightthickness=1).pack(side=tk.LEFT)

        # Send button + clear (far right)
        btn_row = tk.Frame(hdr, bg=C_BG2)
        btn_row.pack(side=tk.RIGHT, padx=8, pady=4)

        self._outline_btn(
            btn_row, "清空",
            lambda: self.tx_text.delete("1.0", tk.END),
            C_TEXT2
        ).pack(side=tk.RIGHT, padx=(4, 0))

        self.send_btn = tk.Button(
            btn_row, text="发  送",
            command=self._send_data,
            font=("Microsoft YaHei", 9, "bold"),
            bg=C_BLUE, fg=C_WHITE,
            activebackground=C_BLUE_HV, activeforeground=C_WHITE,
            relief=tk.FLAT, cursor="hand2",
            padx=16, pady=3, bd=0,
            state=tk.DISABLED
        )
        self.send_btn.pack(side=tk.RIGHT)

        # Hint
        tk.Label(hdr, text="Ctrl+Enter 发送",
                 font=("Microsoft YaHei", 7), fg=C_TEXT2, bg=C_BG2
                 ).pack(side=tk.LEFT, padx=8)

        # Text input
        self.tx_text = tk.Text(
            card, height=4,
            font=("Consolas", 10),
            bg=C_CARD, fg=C_TEXT,
            insertbackground=C_TEXT,
            relief=tk.FLAT, wrap=tk.WORD,
            padx=8, pady=6,
            bd=0
        )
        self.tx_text.pack(fill=tk.X)
        self.tx_text.bind("<Control-Return>", lambda e: self._send_data())

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg=C_BG2, height=26,
                       highlightbackground=C_BORDER, highlightthickness=1)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)

        self.statusbar_var = tk.StringVar(value="就绪 — 请选择串口并打开连接")
        tk.Label(bar, textvariable=self.statusbar_var,
                 font=("Microsoft YaHei", 8), fg=C_TEXT2, bg=C_BG2,
                 anchor=tk.W).pack(side=tk.LEFT, padx=12)

        self.counter_var = tk.StringVar(value="TX: 0 B   RX: 0 B")
        tk.Label(bar, textvariable=self.counter_var,
                 font=("Microsoft YaHei", 8), fg=C_TEXT2, bg=C_BG2
                 ).pack(side=tk.RIGHT, padx=12)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _divider(self, parent):
        tk.Frame(parent, bg=C_BORDER, height=1).pack(fill=tk.X)

    def _outline_btn(self, parent, text, cmd, color):
        return tk.Button(
            parent, text=text, command=cmd,
            font=("Microsoft YaHei", 8),
            bg=C_CARD, fg=color,
            activebackground=C_BG2, activeforeground=color,
            relief=tk.FLAT, cursor="hand2",
            padx=10, pady=3, bd=1,
            highlightbackground=C_BORDER,
            highlightthickness=1
        )

    # ── Serial Operations ─────────────────────────────────────────────────────

    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_combo["values"] = ports
        if ports:
            if self.port_var.get() not in ports:
                self.port_var.set(ports[0])
        else:
            self.port_var.set("")

    def _toggle_connection(self):
        if self.is_connected:
            self._disconnect()
        else:
            self._connect()

    def _connect(self):
        port = self.port_var.get()
        if not port:
            messagebox.showerror("错误", "请先选择串口")
            return

        parity_map = {"None": serial.PARITY_NONE, "Even": serial.PARITY_EVEN,
                      "Odd": serial.PARITY_ODD,  "Mark": serial.PARITY_MARK,
                      "Space": serial.PARITY_SPACE}
        stopbits_map = {"1": serial.STOPBITS_ONE,
                        "1.5": serial.STOPBITS_ONE_POINT_FIVE,
                        "2": serial.STOPBITS_TWO}
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=int(self.baud_var.get()),
                bytesize=int(self.databits_var.get()),
                parity=parity_map[self.parity_var.get()],
                stopbits=stopbits_map[self.stopbits_var.get()],
                timeout=0.1
            )
            self.is_connected = True
            self.running = True
            self.connect_btn.config(text="关闭串口",
                                    bg=C_RED, activebackground="#c0392b")
            self.send_btn.config(state=tk.NORMAL)
            info = (f"{port}  {self.baud_var.get()}bps  "
                    f"{self.databits_var.get()}{self.parity_var.get()[0]}"
                    f"{self.stopbits_var.get()}")
            self.statusbar_var.set(f"已连接: {info}")
            self.status_badge.config(text=f"● {port} 已连接", fg=C_GREEN)
            self._log_sys(f"串口 {port} 已打开")

            self.receive_thread = threading.Thread(
                target=self._receive_loop, daemon=True)
            self.receive_thread.start()
        except serial.SerialException as e:
            messagebox.showerror("连接失败", str(e))

    def _disconnect(self):
        self.running = False
        if self._auto_send_job:
            self.root.after_cancel(self._auto_send_job)
            self._auto_send_job = None
            self.auto_send_var.set(False)

        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()

        self.is_connected = False
        self.connect_btn.config(text="打开串口",
                                bg=C_BLUE, activebackground=C_BLUE_HV)
        self.send_btn.config(state=tk.DISABLED)
        self.statusbar_var.set("连接已关闭")
        self.status_badge.config(text="● 未连接", fg=C_TEXT2)
        self._log_sys("串口已关闭")

    def _receive_loop(self):
        while self.running:
            try:
                if self.serial_port and self.serial_port.is_open:
                    waiting = self.serial_port.in_waiting
                    if waiting > 0:
                        data = self.serial_port.read(waiting)
                        self.rx_count += len(data)
                        self.root.after(0, self._display_rx, data)
                    else:
                        time.sleep(0.01)
            except Exception:
                break

    def _display_rx(self, data: bytes):
        self.rx_text.config(state=tk.NORMAL)

        if self.rx_timestamp_var.get():
            ts = datetime.now().strftime("[%H:%M:%S.%f")[:-3] + "] "
            self.rx_text.insert(tk.END, ts, "ts")

        if self.rx_hex_var.get():
            text = " ".join(f"{b:02X}" for b in data)
            self.rx_text.insert(tk.END, text + "\n", "rx")
        else:
            self._rx_buf.extend(data)
            text = self._flush_rx_buf(self.encoding_var.get())
            if text:
                self.rx_text.insert(tk.END, text, "rx")

        if self.auto_scroll_var.get():
            self.rx_text.see(tk.END)

        self.rx_text.config(state=tk.DISABLED)
        self._update_counters()

    def _flush_rx_buf(self, enc: str) -> str:
        buf = self._rx_buf
        result = []
        i = 0
        while i < len(buf):
            for length in range(1, 5):
                chunk = buf[i:i + length]
                if len(chunk) < length:
                    self._rx_buf = bytearray(buf[i:])
                    return "".join(result)
                try:
                    ch = chunk.decode(enc)
                    result.append(ch)
                    i += length
                    break
                except UnicodeDecodeError:
                    if length == 4:
                        result.append("\ufffd")
                        i += 1
        self._rx_buf = bytearray()
        return "".join(result)

    def _send_data(self):
        if not self.is_connected:
            return
        raw = self.tx_text.get("1.0", tk.END).rstrip("\n")
        if not raw:
            return

        try:
            if self.tx_hex_var.get():
                data = bytes.fromhex(raw.replace(" ", ""))
            else:
                data = raw.encode(self.encoding_var.get())
                if self.tx_newline_var.get():
                    data += b"\r\n"

            self.serial_port.write(data)
            self.tx_count += len(data)

            self.rx_text.config(state=tk.NORMAL)
            if self.rx_timestamp_var.get():
                ts = datetime.now().strftime("[%H:%M:%S.%f")[:-3] + "] "
                self.rx_text.insert(tk.END, ts, "ts")
            echo = (" ".join(f"{b:02X}" for b in data)
                    if self.rx_hex_var.get()
                    else data.decode(self.encoding_var.get(), errors="replace"))
            self.rx_text.insert(tk.END, f"[TX] {echo}\n", "tx")
            if self.auto_scroll_var.get():
                self.rx_text.see(tk.END)
            self.rx_text.config(state=tk.DISABLED)
            self._update_counters()

        except ValueError:
            messagebox.showerror("格式错误", "HEX 格式不正确，示例: FF 0A 1B")
        except Exception as e:
            messagebox.showerror("发送失败", str(e))

    def _toggle_auto_send(self):
        if self.auto_send_var.get():
            self._schedule_auto_send()
        else:
            if self._auto_send_job:
                self.root.after_cancel(self._auto_send_job)
                self._auto_send_job = None

    def _schedule_auto_send(self):
        if self.auto_send_var.get() and self.is_connected:
            self._send_data()
            try:
                interval = max(10, int(self.auto_interval_var.get()))
            except ValueError:
                interval = 1000
            self._auto_send_job = self.root.after(interval, self._schedule_auto_send)

    def _log_sys(self, msg: str):
        self.rx_text.config(state=tk.NORMAL)
        ts = datetime.now().strftime("%H:%M:%S")
        self.rx_text.insert(tk.END, f"[{ts}] {msg}\n", "sys")
        if self.auto_scroll_var.get():
            self.rx_text.see(tk.END)
        self.rx_text.config(state=tk.DISABLED)

    def _clear_rx(self):
        self.rx_text.config(state=tk.NORMAL)
        self.rx_text.delete("1.0", tk.END)
        self.rx_text.config(state=tk.DISABLED)
        self.rx_count = 0
        self._rx_buf = bytearray()
        self._update_counters()

    def _save_rx(self):
        from tkinter.filedialog import asksaveasfilename
        path = asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            title="保存接收数据"
        )
        if path:
            content = self.rx_text.get("1.0", tk.END)
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)

    def _update_counters(self):
        self.counter_var.set(f"TX: {self.tx_count} B   RX: {self.rx_count} B")

    def on_close(self):
        if self.is_connected:
            self._disconnect()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = SerialAssistant(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
