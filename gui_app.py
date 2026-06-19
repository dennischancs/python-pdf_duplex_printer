#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF双面打印延迟控制脚本 - GUI 入口 (tkinter)
解决Brother MFC-7480D等打印机连续双面打印卡纸问题

界面功能:
  - PDF文件/文件夹选择（支持拖拽）
  - 打印机下拉选择（支持刷新）
  - 延迟秒数设置（默认20秒，单面自动禁用）
  - 双面打印模式选择（长边/短边/单面）
  - 打印份数设置（逐份打印）
  - 页面范围选择（全部/奇数/偶数/自定义）— 文件夹时禁用
  - 打印引擎选择
  - 实时进度条和日志显示
  - 后台线程打印，界面不卡顿
  - 支持取消打印

依赖: tkinterdnd2 (用于拖拽支持)
"""

__version__ = "v0.4"


import os
import queue
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

from pdf_duplex_printer import (
    get_available_printers,
    get_default_printer,
    show_printer_capabilities,
    print_with_delay,
    discover_pdf_files,
    list_print_jobs,
    cancel_print_job,
    cancel_all_jobs,
    preview_pdf,
    get_pdf_info,
    find_sumatra_pdf,
    open_system_printer_queue,
)

# 应用内嵌 PDF 预览（依赖 PyMuPDF，不可用时降级为外部打开）
try:
    from pdf_preview import PDFPreviewDialog
    HAS_PREVIEW_DIALOG = True
except ImportError:
    HAS_PREVIEW_DIALOG = False
    PDFPreviewDialog = None

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = None
    DND_FILES = None


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
        self.root.geometry("700x750")
        self.root.minsize(680, 650)

        # 状态变量
        self.print_thread = None
        self.cancel_event = threading.Event()
        self.message_queue = queue.Queue()
        self.total_sheets = 0
        self.is_folder_mode = False  # 是否选择了文件夹

        # 构建 UI
        self._build_ui()

        # 绑定联动事件
        self.duplex_var.trace_add("write", self._on_duplex_change)
        self.page_range_var.trace_add("write", self._on_page_range_change)

        # 初始化控件状态
        self._on_duplex_change()
        self._on_page_range_change()

        # 加载打印机列表
        self._load_printers()

        # 启动消息队列轮询
        self._poll_queue()

        # 设置拖拽支持
        self._setup_drag_drop()

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
        main_frame.rowconfigure(6, weight=1)  # 日志区域可扩展

        # --- 文件/文件夹选择 ---
        file_frame = ttk.LabelFrame(main_frame, text="PDF 文件 / 文件夹", padding="8")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        file_frame.columnconfigure(0, weight=1)

        self.pdf_path = tk.StringVar()
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, state="readonly")
        pdf_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        Tooltip(pdf_entry, "显示已选择的PDF文件或文件夹路径\n"
                       "支持拖拽文件/文件夹到此窗口\n"
                       "或点击右侧按钮选择")

        browse_btn = ttk.Button(file_frame, text="浏览文件", command=self.on_browse, width=10)
        browse_btn.grid(row=0, column=1, padx=(0, 4))
        Tooltip(browse_btn, "选择要打印的PDF文件")

        folder_btn = ttk.Button(file_frame, text="浏览文件夹", command=self.on_browse_folder, width=10)
        folder_btn.grid(row=0, column=2)
        Tooltip(folder_btn, "选择包含PDF文件的文件夹\n将扫描其中所有PDF并依次打印")

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

        info_btn = ttk.Button(printer_frame, text="查看信息", command=self.on_show_info)
        info_btn.grid(row=0, column=2, padx=(4, 0))
        Tooltip(info_btn, "查看选中打印机的当前配置\n（双面模式、纸张大小、方向）")

        # --- 打印份数 ---
        copies_frame = ttk.LabelFrame(main_frame, text="打印份数", padding="8")
        copies_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        self.copies_var = tk.IntVar(value=1)
        copies_label = ttk.Label(copies_frame, text="份数:")
        copies_label.pack(side="left", padx=(0, 4))

        self.copies_spin = ttk.Spinbox(
            copies_frame, from_=1, to=99, textvariable=self.copies_var, width=8
        )
        self.copies_spin.pack(side="left")
        Tooltip(self.copies_spin, "打印份数（默认1份）\n多份时采用逐份打印方式：\n先打完完整一份再打下一份")

        collate_label = ttk.Label(copies_frame, text="（逐份打印）", foreground="#666666")
        collate_label.pack(side="left", padx=(8, 0))

        # --- 打印设置 ---
        settings_frame = ttk.LabelFrame(main_frame, text="打印设置", padding="8")
        settings_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        settings_frame.columnconfigure(1, weight=1)

        # 延迟秒数
        self.delay_label = ttk.Label(settings_frame, text="延迟间隔（秒）:")
        self.delay_label.grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self.delay_var = tk.IntVar(value=20)
        self.delay_spin = ttk.Spinbox(
            settings_frame, from_=1, to=120, textvariable=self.delay_var, width=8
        )
        self.delay_spin.grid(row=0, column=1, sticky="w")
        self.delay_tooltip = Tooltip(self.delay_spin, "每批打印后的等待秒数，让打印机机械结构复位\n"
                            "避免连续双面打印卡纸\n"
                            "Brother MFC-7480D 建议 15-20 秒（默认20秒）")
        # 延迟提示文字（单面模式下显示）
        self.delay_hint = ttk.Label(settings_frame, text="", foreground="#888888")
        self.delay_hint.grid(row=0, column=2, sticky="w", padx=(8, 0))

        # 双面模式
        ttk.Label(settings_frame, text="双面模式:").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        duplex_frame = ttk.Frame(settings_frame)
        duplex_frame.grid(row=1, column=1, sticky="w", pady=(8, 0), columnspan=2)

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

        # 页面范围
        ttk.Label(settings_frame, text="页面范围:").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        pr_frame = ttk.Frame(settings_frame)
        pr_frame.grid(row=2, column=1, sticky="w", pady=(8, 0), columnspan=2)

        self.page_range_var = tk.StringVar(value="all")
        self.pr_all_rb = ttk.Radiobutton(
            pr_frame, text="全部", variable=self.page_range_var, value="all",
            width=6
        )
        self.pr_all_rb.pack(side="left")
        Tooltip(self.pr_all_rb, "打印全部页面")

        self.pr_odd_rb = ttk.Radiobutton(
            pr_frame, text="奇数", variable=self.page_range_var, value="odd",
            width=6
        )
        self.pr_odd_rb.pack(side="left", padx=(4, 0))
        Tooltip(self.pr_odd_rb, "仅打印奇数页（第1,3,5...页）\n适用于手动双面打印场景")

        self.pr_even_rb = ttk.Radiobutton(
            pr_frame, text="偶数", variable=self.page_range_var, value="even",
            width=6
        )
        self.pr_even_rb.pack(side="left", padx=(4, 0))
        Tooltip(self.pr_even_rb, "仅打印偶数页（第2,4,6...页）\n适用于手动双面打印场景")

        self.pr_custom_rb = ttk.Radiobutton(
            pr_frame, text="自定义:", variable=self.page_range_var, value="custom",
        )
        self.pr_custom_rb.pack(side="left", padx=(8, 4))

        self.page_range_custom_entry = ttk.Entry(pr_frame, width=14)
        self.page_range_custom_entry.pack(side="left")
        self.page_range_custom_entry.insert(0, "1-5,8,10-12")
        Tooltip(self.page_range_custom_entry, "自定义页面范围，格式:\n"
                  "- \"1-5\" 表示第1到第5页\n"
                  "- \"1-5,8,10-12\" 表示第1-5、8、10-12页\n"
                  "- 支持逗号分隔和连字符范围\n"
                  "注意: 选择文件夹时此选项不可用")

        # 打印引擎
        ttk.Label(settings_frame, text="打印引擎:").grid(
            row=3, column=0, sticky="w", padx=(0, 8), pady=(8, 0)
        )
        self.engine_var = tk.StringVar(value="auto")
        engine_combo = ttk.Combobox(
            settings_frame, textvariable=self.engine_var,
            values=["auto", "sumatra", "acrobat", "shell"],
            state="readonly", width=15,
        )
        engine_combo.grid(row=3, column=1, sticky="w", pady=(8, 0), columnspan=2)
        Tooltip(engine_combo, "打印引擎选择：\n"
                              "auto - 自动选择（推荐，优先SumatraPDF）\n"
                              "sumatra - 强制使用 SumatraPDF\n"
                              "acrobat - 强制使用 Acrobat Reader\n"
                              "shell - 系统默认程序（降级方案）")

        # 选项
        options_frame = ttk.Frame(settings_frame)
        options_frame.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.keep_temp_var = tk.BooleanVar(value=False)
        self.keep_temp_cb = ttk.Checkbutton(
            options_frame, text="保留临时文件（调试）", variable=self.keep_temp_var
        )
        self.keep_temp_cb.pack(side="left", padx=(0, 16))
        Tooltip(self.keep_temp_cb, "保留拆分后的临时PDF文件\n用于调试和排查问题，普通使用无需勾选")

        self.no_config_var = tk.BooleanVar(value=False)
        self.no_config_cb = ttk.Checkbutton(
            options_frame, text="跳过双面配置", variable=self.no_config_var
        )
        self.no_config_cb.pack(side="left")
        Tooltip(self.no_config_cb, "跳过自动DEVMODE双面配置，使用打印机当前设置\n"
                                   "SumatraPDF引擎下此选项无效（双面通过打印参数控制）")

        # --- 操作按钮 ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        self.print_btn = ttk.Button(btn_frame, text="开始打印", command=self.on_print)
        self.print_btn.grid(row=0, column=0, padx=(0, 8))
        Tooltip(self.print_btn, "开始分批打印PDF文件")

        self.cancel_btn = ttk.Button(
            btn_frame, text="取消打印", command=self.on_cancel, state="disabled"
        )
        self.cancel_btn.grid(row=0, column=1, padx=(0, 8))
        Tooltip(self.cancel_btn, "取消当前正在执行的打印任务\n已发送的打印作业仍会完成")

        self.preview_btn = ttk.Button(btn_frame, text="预览", command=self.on_preview)
        self.preview_btn.grid(row=0, column=2, padx=(0, 8))
        Tooltip(self.preview_btn, "应用内预览PDF，反映双面模式和页面范围设置\n显示装订边标记，不发送打印任务")

        self.queue_btn = ttk.Button(btn_frame, text="打印队列", command=self.on_show_queue)
        self.queue_btn.grid(row=0, column=3)
        Tooltip(self.queue_btn, "查看并管理打印机当前的打印队列\n可取消卡住的打印作业")

        # --- 进度条 ---
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(
            progress_frame, mode="determinate", maximum=100
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.progress_label = ttk.Label(progress_frame, text="就绪")
        self.progress_label.grid(row=0, column=1)

        # --- 日志区域 ---
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="4")
        log_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

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
    # 联动逻辑
    # ============================================================

    def _on_duplex_change(self, *args):
        """双面模式切换时联动延迟控件状态"""
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
                "Brother MFC-7480D 建议 15-20 秒（默认20秒）"
            )

    def _on_page_range_change(self, *args):
        """页面范围切换时联动自定义输入框状态"""
        is_custom = self.page_range_var.get() == "custom"
        if is_custom:
            self.page_range_custom_entry.config(state="normal")
        else:
            self.page_range_custom_entry.config(state="disabled")

    def _set_page_range_controls(self, enabled: bool):
        """启用/禁用页面范围控件（文件夹模式时禁用）"""
        state = "normal" if enabled else "disabled"

        # 始终保持"全部"可选
        self.pr_all_rb.config(state="normal")
        self.pr_odd_rb.config(state=state)
        self.pr_even_rb.config(state=state)
        self.pr_custom_rb.config(state=state)

        if enabled and self.page_range_var.get() == "custom":
            self.page_range_custom_entry.config(state="normal")
        else:
            self.page_range_custom_entry.config(state="disabled")

        if not enabled:
            self.page_range_var.set("all")

    def _get_effective_page_range(self) -> str:
        """获取实际生效的页面范围字符串"""
        if self.is_folder_mode:
            return "all"
        pr = self.page_range_var.get()
        if pr == "custom":
            custom_val = self.page_range_custom_entry.get().strip()
            return custom_val if custom_val else "all"
        return pr

    # ============================================================
    # 事件处理
    # ============================================================

    def on_browse(self):
        """浏览选择PDF文件"""
        filepath = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        )
        if filepath:
            self.pdf_path.set(filepath)
            self.is_folder_mode = False
            self._set_page_range_controls(enabled=True)
            self._log(f"已选择文件: {filepath}")

    def on_browse_folder(self):
        """浏览选择文件夹"""
        folderpath = filedialog.askdirectory(title="选择包含PDF的文件夹")
        if folderpath:
            self.pdf_path.set(folderpath)
            self.is_folder_mode = True
            self._set_page_range_controls(enabled=False)  # 文件夹模式禁用页面范围
            # 统计PDF数量
            pdfs = discover_pdf_files(folderpath)
            self._log(f"已选择文件夹: {folderpath} ({len(pdfs)} 个PDF文件)")

    def on_show_info(self):
        """显示当前选中打印机的信息"""
        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showwarning("提示", "请先选择打印机")
            return

        info = show_printer_capabilities(printer_name)
        model_info = info.get("model_info", {})
        network_text = "是（网络打印机，延迟会自动增加）" if info.get("is_network") else "否（本地打印机）"

        info_text = (
            f"打印机: {info['name']}\n"
            f"双面打印: {info['duplex']}\n"
            f"纸张大小: {info['paper_size']}\n"
            f"方向:     {info['orientation']}\n"
            f"网络打印: {network_text}\n"
        )

        if model_info.get("matched"):
            info_text += (
                f"\n【型号识别】\n"
                f"品牌/型号: {model_info['brand'].upper()} {model_info['model']}\n"
                f"推荐延迟: {model_info['recommended_delay']}秒\n"
                f"备注: {model_info['note']}\n"
                f"数据来源: {'已验证（实测）' if model_info.get('verified') else '通用建议'}\n"
            )
        else:
            info_text += (
                f"\n【型号识别】\n"
                f"推荐延迟: {model_info.get('recommended_delay', '未知')}秒\n"
                f"备注: {model_info.get('note', '未匹配到已知型号')}\n"
            )

        if "error" in info:
            info_text += f"\n错误: {info['error']}"

        messagebox.showinfo("打印机信息", info_text)

    def on_print(self):
        """开始打印"""
        # 验证输入
        pdf_input = self.pdf_path.get()
        if not pdf_input:
            messagebox.showwarning("提示", "请先选择PDF文件或文件夹")
            return
        if not os.path.exists(pdf_input):
            messagebox.showerror("错误", f"不存在:\n{pdf_input}")
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

        try:
            copies = int(self.copies_var.get())
            if copies < 1:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showwarning("提示", "份数必须是大于0的整数")
            return

        # 发现PDF文件
        pdf_files = discover_pdf_files(pdf_input)
        if not pdf_files:
            messagebox.showerror("错误", "未找到PDF文件")
            return

        # 构建确认信息
        mode_names = {"long": "长边翻转", "short": "短边翻转", "none": "单面"}
        engine_names = {"auto": "自动选择", "sumatra": "SumatraPDF",
                        "acrobat": "Acrobat Reader", "shell": "系统默认"}
        duplex_mode = self.duplex_var.get()
        delay_display = "无需 (单面打印)" if duplex_mode == "none" else f"{delay}秒"

        page_range_display = {
            "all": "全部", "odd": "奇数页", "even": "偶数页",
        }.get(self._get_effective_page_range(), self._get_effective_page_range())

        file_info = ""
        if len(pdf_files) == 1:
            file_info = os.path.basename(pdf_files[0])
        else:
            file_info = f"{os.path.basename(pdf_input)} ({len(pdf_files)} 个PDF)"

        confirm = messagebox.askyesno(
            "确认打印",
            f"{'='*40}\n"
            f"  PDF:     {file_info}\n"
            f"  打印机:   {printer_name}\n"
            f"  延迟:     {delay_display}\n"
            f"  双面:     {mode_names[duplex_mode]}\n"
            f"  份数:     {copies} 份（逐份打印）\n"
            f"  页面范围: {page_range_display}"
            f"\n  引擎:     {engine_names.get(self.engine_var.get(), self.engine_var.get())}\n"
            f"{'='*40}\n\n"
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

    def on_preview(self):
        """预览 PDF 文件（应用内嵌预览，反映打印设置）"""
        pdf_input = self.pdf_path.get()
        if not pdf_input:
            messagebox.showwarning("提示", "请先选择PDF文件")
            return

        if os.path.isdir(pdf_input):
            pdfs = discover_pdf_files(pdf_input)
            if not pdfs:
                messagebox.showwarning("提示", "文件夹中没有PDF文件")
                return
            pdf_file = pdfs[0]
            if len(pdfs) > 1:
                self._log(f"文件夹中有 {len(pdfs)} 个PDF文件，预览第一个")
        else:
            pdf_file = pdf_input

        if not os.path.exists(pdf_file):
            messagebox.showerror("预览失败", f"找不到文件:\n{pdf_file}")
            return

        # 读取当前打印设置
        duplex_mode = self.duplex_var.get()          # long / short / none
        page_range = self.page_range_var.get()        # all / odd / even / custom
        custom_range_str = self.page_range_custom_entry.get() if page_range == "custom" else ""

        self._log(f"打开预览: {os.path.basename(pdf_file)} "
                  f"(双面:{duplex_mode}, 范围:{page_range}"
                  f"{f'[{custom_range_str}]' if custom_range_str else ''})")

        # 优先使用应用内嵌预览（需 PyMuPDF）
        if HAS_PREVIEW_DIALOG:
            try:
                PDFPreviewDialog(self.root, pdf_file, duplex_mode,
                                 page_range, custom_range_str)
                self._log("预览窗口已打开")
            except Exception as ex:
                self._log(f"内嵌预览异常: {ex}，降级为外部打开")
                if not preview_pdf(pdf_file):
                    messagebox.showerror("预览失败",
                        f"无法打开文件:\n{pdf_file}\n\n"
                        f"内嵌预览异常: {ex}\n"
                        f"外部预览也失败，请确认已安装 SumatraPDF。")
        else:
            # PyMuPDF 不可用，降级为外部 SumatraPDF 打开
            if not preview_pdf(pdf_file):
                messagebox.showerror("预览失败",
                    f"无法打开文件:\n{pdf_file}\n\n"
                    f"应用内嵌预览不可用（缺少 PyMuPDF 库），\n"
                    f"外部预览也失败，请确认已安装 SumatraPDF 或系统默认的 PDF 阅读器。")
            else:
                self._log("预览已通过外部程序打开（内嵌预览不可用）")

    def on_show_queue(self):
        """显示打印队列管理对话框"""
        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showwarning("提示", "请先选择打印机")
            return

        # 创建队列对话框
        queue_dialog = tk.Toplevel(self.root)
        queue_dialog.title(f"打印队列 - {printer_name}")
        queue_dialog.geometry("600x440")
        queue_dialog.minsize(550, 340)
        queue_dialog.transient(self.root)
        queue_dialog.grab_set()

        # 树形表格
        columns = ("job_id", "document", "status", "pages", "owner")
        tree = ttk.Treeview(queue_dialog, columns=columns, show="headings", selectmode="browse")
        tree.heading("job_id", text="作业ID")
        tree.heading("document", text="文档名")
        tree.heading("status", text="状态")
        tree.heading("pages", text="页数")
        tree.heading("owner", text="提交者")

        tree.column("job_id", width=60, anchor="center")
        tree.column("document", width=200)
        tree.column("status", width=90, anchor="center")
        tree.column("pages", width=50, anchor="center")
        tree.column("owner", width=100)

        scrollbar = ttk.Scrollbar(queue_dialog, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.grid(row=0, column=0, columnspan=4, sticky="nsew", padx=8, pady=(8, 4))
        scrollbar.grid(row=0, column=4, sticky="ns", pady=(8, 4))

        queue_dialog.columnconfigure(0, weight=1)
        queue_dialog.rowconfigure(0, weight=1)

        # 按钮
        btn_frame = ttk.Frame(queue_dialog)
        btn_frame.grid(row=1, column=0, columnspan=4, sticky="ew", padx=8, pady=4)

        # 状态标签（显示诊断信息）
        status_label = ttk.Label(queue_dialog, text="正在加载...", anchor="center")
        status_label.grid(row=2, column=0, columnspan=4, sticky="ew", padx=8, pady=(0, 4))

        def refresh_queue():
            for item in tree.get_children():
                tree.delete(item)
            try:
                result = list_print_jobs(printer_name)
            except Exception as ex:
                status_label.configure(text=f"获取失败: {ex}")
                return
            jobs = result.get("jobs", [])
            method = result.get("method", "none")
            diagnostics = result.get("diagnostics", "")
            error = result.get("error", "")
            if not jobs:
                if error:
                    status_label.configure(
                        text=f"(队列为空) {diagnostics}  [错误: {error[:60]}]",
                        foreground="#C62828")
                else:
                    status_label.configure(text=f"(队列为空) {diagnostics}",
                                           foreground="#666666")
            else:
                method_label = {"EnumJobs-L1": "API", "EnumJobs-L2": "API(L2)",
                                "WMI": "WMI"}.get(method, method)
                status_label.configure(
                    text=f"共 {len(jobs)} 个待处理作业  [获取方式: {method_label}]",
                    foreground="#2E7D32")
            for job in jobs:
                tree.insert("", "end", values=(
                    job["job_id"], job["document"], job["status"],
                    job["pages"], job["owner"],
                ))

        refresh_btn = ttk.Button(btn_frame, text="刷新", command=refresh_queue)
        refresh_btn.pack(side="left", padx=(0, 8))

        def cancel_selected():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("提示", "请先选择要取消的作业")
                return
            item = selection[0]
            values = tree.item(item, "values")
            job_id = int(values[0])
            doc = values[1]

            confirm = messagebox.askyesno(
                "确认取消", f"确定要取消作业 #{job_id} 吗？\n\n文档: {doc}"
            )
            if confirm:
                if cancel_print_job(printer_name, job_id):
                    self._log(f"已取消作业 #{job_id}: {doc}")
                    refresh_queue()
                else:
                    messagebox.showerror("错误", f"取消失败: 作业 #{job_id}")

        cancel_btn = ttk.Button(btn_frame, text="取消选中作业", command=cancel_selected)
        cancel_btn.pack(side="left", padx=(0, 8))

        def cancel_all():
            confirm = messagebox.askyesno(
                "确认取消全部", f"确定要取消 '{printer_name}' 的所有打印作业吗？\n\n此操作不可撤销！"
            )
            if confirm:
                cancelled, failed = cancel_all_jobs(printer_name)
                self._log(f"已取消 {cancelled} 个作业（{failed} 个失败）")
                refresh_queue()
                messagebox.showinfo("完成", f"已取消 {cancelled} 个作业")

        cancel_all_btn = ttk.Button(btn_frame, text="取消全部", command=cancel_all)
        cancel_all_btn.pack(side="left", padx=(0, 8))

        def open_sys_queue():
            """打开 Windows 系统打印队列窗口（25H2 兼容兜底）"""
            if open_system_printer_queue(printer_name):
                self._log(f"已打开系统打印队列窗口: {printer_name}")
            else:
                messagebox.showerror("错误",
                    f"无法打开系统打印队列窗口。\n"
                    f"打印机: {printer_name}\n\n"
                    f"可手动打开: 设置 → 蓝牙和设备 → 打印机和扫描仪 → 选择打印机 → 打开队列")

        sys_queue_btn = ttk.Button(btn_frame, text="打开系统队列", command=open_sys_queue)
        sys_queue_btn.pack(side="left", padx=(0, 8))
        Tooltip(sys_queue_btn, "打开 Windows 系统自带的打印队列窗口\n"
                               "当应用内队列无法显示作业时，可使用此功能作为兜底\n"
                               "（Windows 11 25H2 兼容）")

        close_btn = ttk.Button(btn_frame, text="关闭", command=queue_dialog.destroy)
        close_btn.pack(side="right")

        # 初始加载
        refresh_queue()

    # ============================================================
    # 拖拽支持
    # ============================================================

    def _setup_drag_drop(self):
        """设置拖拽支持（需要 tkinterdnd2 库）"""
        if not HAS_DND:
            self._log("提示: 安装 tkinterdnd2 可启用拖拽支持 (pip install tkinterdnd2)")
            return

        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._on_drop)
        except Exception as e:
            self._log(f"拖拽初始化失败: {e}")

    def _parse_drop_paths(self, data: str) -> list:
        """
        解析 tkinterdnd2 的拖拽数据
        处理带空格路径的 {} 包裹格式
        """
        paths = []
        # 匹配 {} 包裹的路径（带空格） 或 不带{} 的路径（无空格）
        for match in re.findall(r'\{([^}]+)\}|(\S+)', data):
            path = match[0] or match[1]
            if path:
                paths.append(path)
        return paths

    def _on_drop(self, event):
        """处理拖拽文件/文件夹事件"""
        paths = self._parse_drop_paths(event.data)
        if not paths:
            return

        path = paths[0]

        if os.path.isdir(path):
            self.pdf_path.set(path)
            self.is_folder_mode = True
            self._set_page_range_controls(enabled=False)
            pdfs = discover_pdf_files(path)
            self._log(f"拖入文件夹: {os.path.basename(path)} ({len(pdfs)} 个PDF)")
        elif path.lower().endswith('.pdf') or path.lower().endswith('.pdf"'):
            clean_path = path.rstrip('"')
            self.pdf_path.set(clean_path)
            self.is_folder_mode = False
            self._set_page_range_controls(enabled=True)
            self._log(f"拖入文件: {os.path.basename(clean_path)}")
        else:
            messagebox.showwarning("提示", "只支持 PDF 文件或文件夹")

    # ============================================================
    # 后台打印线程
    # ============================================================

    def _print_worker(self):
        """后台线程：执行打印，通过队列向主线程报告进度"""

        def progress_callback(event_type, data):
            self.message_queue.put((event_type, data))

        try:
            pdf_input = self.pdf_path.get()
            pdf_files = discover_pdf_files(pdf_input)

            if not pdf_files:
                self.message_queue.put(("error", {"message": "未找到PDF文件"}))
                return

            copies = int(self.copies_var.get())
            page_range = self._get_effective_page_range()

            total_result = {"success_count": 0, "total_sheets": 0, "cancelled": False}

            for fidx, pdf_file in enumerate(pdf_files, 1):
                # 报告开始处理新文件
                self.message_queue.put(("file_start", {
                    "index": fidx,
                    "total": len(pdf_files),
                    "filepath": pdf_file,
                }))

                result = print_with_delay(
                    pdf_file=pdf_file,
                    delay_seconds=int(self.delay_var.get()),
                    printer_name=self.printer_var.get(),
                    keep_temp=self.keep_temp_var.get(),
                    configure_duplex=not self.no_config_var.get(),
                    duplex_mode=self.duplex_var.get(),
                    engine=self.engine_var.get(),
                    copies=copies,
                    page_range=page_range,
                    progress_callback=progress_callback,
                    cancel_event=self.cancel_event,
                )

                total_result["success_count"] += result["success_count"]
                total_result["total_sheets"] += result["total_sheets"]

                self.message_queue.put(("file_done", {
                    "index": fidx,
                    "total": len(pdf_files),
                    "success_count": result["success_count"],
                    "total_sheets": result["total_sheets"],
                }))

                if result["cancelled"]:
                    total_result["cancelled"] = True
                    break

            self.message_queue.put(("done", total_result))

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

        elif event_type == "copy_start":
            self._log(f">>> 第 {data['copy_index']}/{data['total_copies']} 份 <<<")

        elif event_type == "file_start":
            fname = os.path.basename(data.get("filepath", ""))
            total = data.get("total", 1)
            idx = data.get("index", 1)
            if total > 1:
                self._log(f"\n[{idx}/{total}] ====== {fname} ======")

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

        elif event_type == "file_done":
            idx = data["index"]
            total = data["total"]
            sc = data["success_count"]
            ts = data["total_sheets"]
            if total > 1:
                self._log(f"[{idx}/{total}] 完成: {sc}/{ts} 批次")

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
            else:
                self._log(f"\n打印任务完成! 成功 {success}/{total} 批次")
                self.progress_label.config(text=f"完成 {success}/{total}")

            self._set_printing_state(False)

        elif event_type == "stats":
            # 打印统计
            self._log(f"\n{'='*50}")
            self._log(f"打印统计")
            self._log(f"{'='*50}")
            self._log(f"  总页数:       {data['total_pages']}")
            self._log(f"  总批次数:     {data['total_batches']}")
            self._log(f"  成功批次:     {data['success_batches']}")
            self._log(f"  失败批次:     {data['failed_batches']}")
            self._log(f"  成功率:       {data['success_rate']:.1f}%")
            self._log(f"  总耗时:       {data['elapsed_formatted']}")
            self._log(f"{'='*50}")

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
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app = DuplexPrinterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
