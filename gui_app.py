#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF双面打印延迟控制脚本 - GUI 入口 (tkinter)
解决Brother MFC-7480D等打印机连续双面打印卡纸问题

界面功能:
  - PDF文件选择
  - 打印机下拉选择（支持刷新）
  - 延迟秒数设置
  - 双面打印模式选择（长边/短边/单面）
  - 实时进度条和日志显示
  - 后台线程打印，界面不卡顿
  - 支持取消打印
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

from pdf_duplex_printer import (
    get_available_printers,
    get_default_printer,
    show_printer_capabilities,
    print_with_delay,
)


class Tooltip:
    """
    Tkinter 控件悬停提示（纯 tkinter 实现，无外部依赖）
    鼠标悬停 500ms 后显示提示，离开或点击时隐藏
    """

    def __init__(self, widget, text: str, delay: int = 500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip_window = None
        self.after_id = None

        widget.bind("<Enter>", self._on_enter)
        widget.bind("<Leave>", self._on_leave)
        widget.bind("<Button-1>", self._on_leave)

    def _on_enter(self, event=None):
        """鼠标进入，安排延迟显示"""
        self._cancel()
        self.after_id = self.widget.after(self.delay, self._show)

    def _on_leave(self, event=None):
        """鼠标离开或点击，隐藏提示"""
        self._cancel()
        self._hide()

    def _cancel(self):
        """取消延迟显示"""
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def _show(self):
        """显示提示窗口"""
        if self.tooltip_window or not self.text:
            return

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)  # 无边框
        self.tooltip_window.wm_attributes("-topmost", True)  # 置顶

        label = tk.Label(
            self.tooltip_window,
            text=self.text,
            justify=tk.LEFT,
            background="#FFFFE0",  # 浅黄色背景
            foreground="#333333",
            relief=tk.SOLID,
            borderwidth=1,
            font=("Microsoft YaHei", 9),
            padx=6,
            pady=3,
        )
        label.pack()

        # 定位：鼠标位置右下方
        x = self.widget.winfo_pointerx() + 15
        y = self.widget.winfo_pointery() + 10

        # 屏幕边缘修正
        screen_w = self.widget.winfo_screenwidth()
        screen_h = self.widget.winfo_screenheight()
        win_w = self.tooltip_window.winfo_reqwidth()
        win_h = self.tooltip_window.winfo_reqheight()
        if x + win_w > screen_w:
            x = screen_w - win_w - 5
        if y + win_h > screen_h:
            y = self.widget.winfo_pointery() - win_h - 10

        self.tooltip_window.wm_geometry(f"+{x}+{y}")

    def _hide(self):
        """隐藏提示窗口"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class DuplexPrinterGUI:
    """PDF双面打印延迟控制 GUI 主类"""

    def __init__(self, root):
        self.root = root
        self.root.title("PDF 双面打印延迟控制")
        self.root.geometry("620x580")
        self.root.minsize(580, 520)

        # 状态变量
        self.print_thread = None
        self.cancel_event = threading.Event()
        self.message_queue = queue.Queue()
        self.total_sheets = 0

        # 构建 UI
        self._build_ui()

        # 绑定双面模式切换事件（联动延迟控件状态）
        self.duplex_var.trace_add("write", self._on_duplex_change)
        # 初始化延迟控件状态
        self._on_duplex_change()

        # 加载打印机列表
        self._load_printers()

        # 启动消息队列轮询
        self._poll_queue()

    # ============================================================
    # UI 构建
    # ============================================================

    def _build_ui(self):
        """构建界面布局"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="12")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # --- 文件选择 ---
        file_frame = ttk.LabelFrame(main_frame, text="PDF 文件", padding="8")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        file_frame.columnconfigure(0, weight=1)

        self.pdf_path = tk.StringVar()
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, state="readonly")
        pdf_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        Tooltip(pdf_entry, "显示已选择的PDF文件路径\n点击右侧「浏览...」按钮选择文件")

        browse_btn = ttk.Button(file_frame, text="浏览...", command=self.on_browse)
        browse_btn.grid(row=0, column=1)
        Tooltip(browse_btn, "点击打开文件选择对话框，选择要打印的PDF文件")

        # --- 打印机选择 ---
        printer_frame = ttk.LabelFrame(main_frame, text="打印机", padding="8")
        printer_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        printer_frame.columnconfigure(0, weight=1)

        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(
            printer_frame, textvariable=self.printer_var, state="readonly"
        )
        self.printer_combo.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        Tooltip(self.printer_combo, "选择目标打印机\n点击「刷新」按钮可重新获取打印机列表")

        refresh_btn = ttk.Button(printer_frame, text="刷新", command=self._load_printers)
        refresh_btn.grid(row=0, column=1)
        Tooltip(refresh_btn, "重新扫描系统中的可用打印机")

        # 打印机信息按钮
        info_btn = ttk.Button(printer_frame, text="查看信息", command=self.on_show_info)
        info_btn.grid(row=0, column=2, padx=(4, 0))
        Tooltip(info_btn, "查看选中打印机的当前配置\n（双面模式、纸张大小、方向）")

        # --- 参数设置 ---
        settings_frame = ttk.LabelFrame(main_frame, text="打印设置", padding="8")
        settings_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        settings_frame.columnconfigure(1, weight=1)

        # 延迟秒数
        self.delay_label = ttk.Label(settings_frame, text="延迟间隔（秒）:")
        self.delay_label.grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self.delay_var = tk.IntVar(value=15)
        self.delay_spin = ttk.Spinbox(
            settings_frame, from_=1, to=120, textvariable=self.delay_var, width=8
        )
        self.delay_spin.grid(row=0, column=1, sticky="w")
        self.delay_tooltip = Tooltip(self.delay_spin, "每批打印后的等待秒数，让打印机机械结构复位\n"
                            "避免连续双面打印卡纸\n"
                            "Brother MFC-7480D 建议 15-20 秒")
        # 延迟提示文字（单面模式下显示）
        self.delay_hint = ttk.Label(settings_frame, text="", foreground="#888888")
        self.delay_hint.grid(row=0, column=2, sticky="w", padx=(8, 0))

        # 双面模式
        ttk.Label(settings_frame, text="双面模式:").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        duplex_frame = ttk.Frame(settings_frame)
        duplex_frame.grid(row=1, column=1, sticky="w", pady=(8, 0))

        self.duplex_var = tk.StringVar(value="long")
        self.duplex_long_rb = ttk.Radiobutton(
            duplex_frame, text="长边翻转", variable=self.duplex_var, value="long"
        )
        self.duplex_long_rb.grid(row=0, column=0, padx=(0, 12))
        Tooltip(self.duplex_long_rb, "长边翻转（书本式）\n沿长边翻页，适合竖向文档")

        self.duplex_short_rb = ttk.Radiobutton(
            duplex_frame, text="短边翻转", variable=self.duplex_var, value="short"
        )
        self.duplex_short_rb.grid(row=0, column=1, padx=(0, 12))
        Tooltip(self.duplex_short_rb, "短边翻转（记事本式）\n沿短边翻页，适合横向文档")

        self.duplex_none_rb = ttk.Radiobutton(
            duplex_frame, text="单面", variable=self.duplex_var, value="none"
        )
        self.duplex_none_rb.grid(row=0, column=2)
        Tooltip(self.duplex_none_rb, "单面打印\n不使用双面功能，只打印一面\n"
                                      "选择此项时延迟间隔将自动禁用（单面无需延迟）")

        # 打印引擎
        ttk.Label(settings_frame, text="打印引擎:").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.engine_var = tk.StringVar(value="auto")
        engine_combo = ttk.Combobox(
            settings_frame, textvariable=self.engine_var,
            values=["auto", "sumatra", "acrobat", "shell"],
            state="readonly", width=15,
        )
        engine_combo.grid(row=2, column=1, sticky="w", pady=(8, 0))
        Tooltip(engine_combo, "打印引擎选择：\n"
                              "auto - 自动选择（推荐，优先SumatraPDF）\n"
                              "sumatra - 强制使用 SumatraPDF\n"
                              "acrobat - 强制使用 Acrobat Reader\n"
                              "shell - 系统默认程序（降级方案）")

        # 选项
        options_frame = ttk.Frame(settings_frame)
        options_frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.keep_temp_var = tk.BooleanVar(value=False)
        self.keep_temp_cb = ttk.Checkbutton(
            options_frame, text="保留临时文件（调试）", variable=self.keep_temp_var
        )
        self.keep_temp_cb.grid(row=0, column=0, padx=(0, 16))
        Tooltip(self.keep_temp_cb, "保留拆分后的临时PDF文件\n用于调试和排查问题，普通使用无需勾选")

        self.no_config_var = tk.BooleanVar(value=False)
        self.no_config_cb = ttk.Checkbutton(
            options_frame, text="跳过双面配置", variable=self.no_config_var
        )
        self.no_config_cb.grid(row=0, column=1)
        Tooltip(self.no_config_cb, "跳过自动DEVMODE双面配置，使用打印机当前设置\n"
                                   "SumatraPDF引擎下此选项无效（双面通过打印参数控制）")

        # --- 操作按钮 ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        self.print_btn = ttk.Button(btn_frame, text="开始打印", command=self.on_print)
        self.print_btn.grid(row=0, column=0, padx=(0, 8))
        Tooltip(self.print_btn, "开始分批打印PDF文件")

        self.cancel_btn = ttk.Button(
            btn_frame, text="取消打印", command=self.on_cancel, state="disabled"
        )
        self.cancel_btn.grid(row=0, column=1, padx=(0, 8))
        Tooltip(self.cancel_btn, "取消当前正在执行的打印任务\n已发送的打印作业仍会完成")

        # --- 进度条 ---
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(
            progress_frame, mode="determinate", maximum=100
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.progress_label = ttk.Label(progress_frame, text="就绪")
        self.progress_label.grid(row=0, column=1)

        # --- 日志区域 ---
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="4")
        log_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=10, state="disabled", wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

    # ============================================================
    # 打印机加载
    # ============================================================

    def _load_printers(self):
        """加载打印机列表到下拉框"""
        try:
            printers = get_available_printers()
            default_printer = get_default_printer()

            self.printer_combo["values"] = printers
            if default_printer in printers:
                self.printer_var.set(default_printer)
            elif printers:
                self.printer_var.set(printers[0])

            self._log(f"已加载 {len(printers)} 台打印机")
        except Exception as e:
            self._log(f"加载打印机失败: {e}")
            messagebox.showerror("错误", f"加载打印机失败:\n{e}")

    # ============================================================
    # 事件处理
    # ============================================================

    def _on_duplex_change(self, *args):
        """双面模式切换时联动延迟控件状态

        单面模式: 禁用延迟Spinbox，显示"单面无需延迟"提示
        双面模式: 启用延迟Spinbox，恢复正常提示
        """
        is_simplex = self.duplex_var.get() == "none"

        if is_simplex:
            self.delay_spin.config(state="disabled")
            self.delay_label.config(foreground="#999999")
            self.delay_hint.config(text="单面无需延迟")
            self.delay_tooltip.text = (
                "单面打印模式下无需设置延迟\n"
                "延迟仅用于双面打印时防止打印机机械结构卡纸\n"
                "切换到双面模式后将自动恢复此设置"
            )
        else:
            self.delay_spin.config(state="normal")
            self.delay_label.config(foreground="#000000")
            self.delay_hint.config(text="")
            self.delay_tooltip.text = (
                "每批打印后的等待秒数，让打印机机械结构复位\n"
                "避免连续双面打印卡纸\n"
                "Brother MFC-7480D 建议 15-20 秒"
            )

    def on_browse(self):
        """浏览选择PDF文件"""
        filepath = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        )
        if filepath:
            self.pdf_path.set(filepath)
            self._log(f"已选择文件: {filepath}")

    def on_show_info(self):
        """显示当前选中打印机的信息"""
        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showwarning("提示", "请先选择打印机")
            return

        info = show_printer_capabilities(printer_name)
        info_text = (
            f"打印机: {info['name']}\n"
            f"双面打印: {info['duplex']}\n"
            f"纸张大小: {info['paper_size']}\n"
            f"方向: {info['orientation']}"
        )
        if "error" in info:
            info_text += f"\n错误: {info['error']}"

        messagebox.showinfo("打印机信息", info_text)

    def on_print(self):
        """开始打印"""
        # 验证输入
        pdf_file = self.pdf_path.get()
        if not pdf_file:
            messagebox.showwarning("提示", "请先选择PDF文件")
            return
        if not os.path.exists(pdf_file):
            messagebox.showerror("错误", f"文件不存在:\n{pdf_file}")
            return

        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showwarning("提示", "请先选择打印机")
            return

        try:
            delay = int(self.delay_var.get())
            if delay < 1:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showwarning("提示", "延迟间隔必须是大于0的整数")
            return

        # 确认开始
        mode_names = {"long": "长边翻转", "short": "短边翻转", "none": "单面"}
        engine_names = {"auto": "自动选择", "sumatra": "SumatraPDF",
                        "acrobat": "Acrobat Reader", "shell": "系统默认"}
        duplex_mode = self.duplex_var.get()
        delay_display = "无需 (单面打印)" if duplex_mode == "none" else f"{delay}秒"
        confirm = messagebox.askyesno(
            "确认打印",
            f"PDF文件: {os.path.basename(pdf_file)}\n"
            f"打印机: {printer_name}\n"
            f"延迟间隔: {delay_display}\n"
            f"双面模式: {mode_names[duplex_mode]}\n"
            f"打印引擎: {engine_names.get(self.engine_var.get(), self.engine_var.get())}\n\n"
            f"确认开始打印？"
        )
        if not confirm:
            return

        # 重置状态
        self.total_sheets = 0
        self.progress_bar["value"] = 0
        self.progress_label.config(text="准备中...")
        self._clear_log()

        # 切换按钮状态
        self._set_printing_state(True)

        # 启动打印线程
        self.cancel_event.clear()
        self.print_thread = threading.Thread(target=self._print_worker, daemon=True)
        self.print_thread.start()

    def on_cancel(self):
        """取消打印"""
        if self.print_thread and self.print_thread.is_alive():
            confirm = messagebox.askyesno("确认取消", "确定要取消当前打印任务吗？")
            if confirm:
                self.cancel_event.set()
                self._log("正在取消打印...")
                self.cancel_btn.config(state="disabled")

    # ============================================================
    # 后台打印线程
    # ============================================================

    def _print_worker(self):
        """后台线程：执行打印，通过队列向主线程报告进度"""

        def progress_callback(event_type, data):
            self.message_queue.put((event_type, data))

        try:
            result = print_with_delay(
                pdf_file=self.pdf_path.get(),
                delay_seconds=int(self.delay_var.get()),
                printer_name=self.printer_var.get(),
                keep_temp=self.keep_temp_var.get(),
                configure_duplex=not self.no_config_var.get(),
                duplex_mode=self.duplex_var.get(),
                engine=self.engine_var.get(),
                progress_callback=progress_callback,
                cancel_event=self.cancel_event,
            )
            self.message_queue.put(("done", result))

        except Exception as e:
            self.message_queue.put(("error", {"message": str(e)}))

    # ============================================================
    # 消息队列轮询（线程安全 UI 更新）
    # ============================================================

    def _poll_queue(self):
        """每100ms检查消息队列，更新UI"""
        try:
            while not self.message_queue.empty():
                event_type, data = self.message_queue.get_nowait()
                self._handle_event(event_type, data)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_queue)

    def _handle_event(self, event_type, data):
        """处理队列消息，更新进度条和日志"""

        if event_type == "info":
            self._log(data.get("message", ""))

        elif event_type == "engine":
            engine_name = data.get("engine_name", "")
            engine_path = data.get("engine_path", "")
            self._log(f"使用打印引擎: {engine_name}")
            if engine_path and engine_path != "N/A":
                self._log(f"  引擎路径: {engine_path}")

        elif event_type == "split_start":
            self.total_sheets = data["sheets"]
            self._log(f"PDF总页数: {data['total_pages']}")
            if data.get("is_duplex", True):
                self._log(f"需要打印 {data['sheets']} 张纸（双面）")
            else:
                self._log(f"单面打印，共 {data['total_pages']} 页直接打印")
            self.progress_bar["maximum"] = data["sheets"]
            self.progress_label.config(text=f"0/{data['sheets']}")

        elif event_type == "split_done":
            self._log(f"PDF拆分完成，共 {data['sheets']} 个文件")

        elif event_type == "config":
            mode_names = {"long": "长边翻转", "short": "短边翻转", "none": "单面"}
            mode = mode_names.get(data["duplex_mode"], data["duplex_mode"])
            if data["success"]:
                self._log(f"  已配置打印机: {mode}")
            else:
                self._log(f"  警告: 打印机配置失败，使用当前设置")

        elif event_type == "batch_start":
            idx = data["index"]
            total = data["total"]
            self._log(f"[{idx}/{total}] 正在打印: {data['filename']}")

        elif event_type == "batch_done":
            idx = data["index"]
            total = data["total"]
            self.progress_bar["value"] = idx
            pct = int(idx / total * 100) if total > 0 else 0
            self.progress_label.config(text=f"{idx}/{total} ({pct}%)")
            self._log(f"  打印命令已发送")

        elif event_type == "batch_fail":
            self._log(f"  警告: 打印失败 - {data.get('error', '')}")

        elif event_type == "delay":
            remaining = data["remaining"]
            total = data["total"]
            self.progress_label.config(text=f"等待中... {remaining}s")

        elif event_type == "cleanup":
            if not data["kept"]:
                self._log("清理临时文件...")

        elif event_type == "done":
            success = data["success_count"]
            total = data["total_sheets"]
            cancelled = data["cancelled"]

            if cancelled:
                self._log(f"\n打印已取消! 已完成 {success}/{total} 批次")
                self.progress_label.config(text=f"已取消 {success}/{total}")
                messagebox.showinfo("打印取消", f"已取消打印\n完成 {success}/{total} 批次")
            else:
                self._log(f"\n打印任务完成! 成功 {success}/{total} 批次")
                self.progress_label.config(text=f"完成 {success}/{total}")
                messagebox.showinfo("打印完成", f"打印任务完成!\n成功 {success}/{total} 批次")

            self._set_printing_state(False)

        elif event_type == "error":
            self._log(f"错误: {data.get('message', '')}")
            messagebox.showerror("错误", data.get("message", "未知错误"))
            self._set_printing_state(False)

    # ============================================================
    # UI 辅助方法
    # ============================================================

    def _set_printing_state(self, printing: bool):
        """切换打印/非打印状态的按钮"""
        if printing:
            self.print_btn.config(state="disabled")
            self.cancel_btn.config(state="normal")
        else:
            self.print_btn.config(state="normal")
            self.cancel_btn.config(state="disabled")

    def _log(self, message):
        """向日志区域追加一行文本"""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _clear_log(self):
        """清空日志区域"""
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")


# ============================================================
# 入口
# ============================================================

def main():
    root = tk.Tk()
    app = DuplexPrinterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
