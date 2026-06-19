# -*- coding: utf-8 -*-
"""
PDF 打印预览模块 (v0.4)

使用 PyMuPDF (fitz) 将 PDF 页面渲染为图像，在 tkinter 窗口内嵌预览。
反映打印设置（双面模式装订边、页面范围筛选），支持翻页与缩放。
PyMuPDF 不可用时自动降级为外部程序打开。

依赖: PyMuPDF (fitz)
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox

# 尝试导入 PyMuPDF
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    try:
        import pymupdf as fitz  # 新版导入名
        HAS_FITZ = True
    except ImportError:
        HAS_FITZ = False
        fitz = None

# 复用核心模块的页面范围解析与降级预览
try:
    from pdf_duplex_printer import parse_page_range, preview_pdf
except ImportError:
    parse_page_range = None
    preview_pdf = None


def render_pdf_page(pdf_path, page_index, dpi=150, zoom=1.0):
    """
    用 PyMuPDF 渲染指定页为 PNG 字节数据。

    参数:
        pdf_path: PDF 文件路径
        page_index: 0-based 页码索引
        dpi: 渲染分辨率（默认150）
        zoom: 额外缩放系数（默认1.0）

    返回: PNG 字节数据，失败返回 None
    """
    if not HAS_FITZ:
        return None
    try:
        doc = fitz.open(pdf_path)
        try:
            if page_index < 0 or page_index >= len(doc):
                return None
            page = doc[page_index]
            # 组合 dpi 与 zoom：zoom 直接作为缩放矩阵
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, dpi=dpi)
            return pix.tobytes("png")
        finally:
            doc.close()
    except Exception:
        return None


def get_pdf_page_count(pdf_path):
    """获取 PDF 总页数，失败返回 0"""
    if not HAS_FITZ:
        return 0
    try:
        doc = fitz.open(pdf_path)
        try:
            return len(doc)
        finally:
            doc.close()
    except Exception:
        return 0


class PDFPreviewDialog:
    """
    应用内嵌 PDF 打印预览对话框。

    在 tkinter Toplevel 中渲染 PDF 页面，反映双面模式装订边与页面范围设置。
    支持翻页、缩放、装订标记开关。渲染在后台线程执行，避免阻塞 UI。
    """

    # 装订边颜色
    BIND_COLOR = "#E53935"
    MARGIN = 24  # Canvas 内边距

    def __init__(self, parent, pdf_path, duplex_mode="long",
                 page_range="all", custom_range_str=""):
        """
        参数:
            parent: 父窗口
            pdf_path: PDF 文件路径
            duplex_mode: 双面模式 "long"(长边)/"short"(短边)/"none"(单面)
            page_range: 页面范围 "all"/"odd"/"even"/"custom"
            custom_range_str: 自定义范围字符串（page_range="custom" 时使用）
        """
        self.parent = parent
        self.pdf_path = pdf_path
        self.duplex_mode = duplex_mode
        self.page_range = page_range
        self.custom_range_str = custom_range_str

        # PyMuPDF 不可用 → 降级
        if not HAS_FITZ:
            self._fallback_external()
            return

        # 计算总页数与预览页码列表
        self.total_pages = get_pdf_page_count(pdf_path)
        if self.total_pages == 0:
            messagebox.showerror("预览失败",
                f"无法读取 PDF 文件，或文件为空:\n{pdf_path}")
            return

        self.page_indices = self._compute_page_indices()
        if not self.page_indices:
            messagebox.showwarning("预览",
                "当前页面范围设置下没有可预览的页面。\n"
                f"总页数: {self.total_pages}")
            return

        # 当前在 page_indices 中的位置
        self.current_pos = 0
        self.zoom = 1.0
        self.show_bind_marker = (duplex_mode in ("long", "short"))

        # 渲染线程控制
        self._render_token = 0
        self._render_lock = threading.Lock()
        self._current_photo = None  # 持有 PhotoImage 引用避免被GC

        self._build_ui()
        self._render_current()

    # ----------------------------------------------------------
    # 降级处理
    # ----------------------------------------------------------
    def _fallback_external(self):
        """PyMuPDF 不可用时，降级为外部程序打开"""
        messagebox.showinfo("预览",
            "应用内嵌预览需要 PyMuPDF 库，当前不可用。\n"
            "将使用外部程序（SumatraPDF）打开预览。")
        if preview_pdf is not None:
            if not preview_pdf(self.pdf_path):
                messagebox.showerror("预览失败",
                    f"无法打开文件:\n{self.pdf_path}\n\n"
                    "请确认已安装 SumatraPDF 或系统默认的 PDF 阅读器。")
        else:
            # 兜底：os.startfile
            try:
                os.startfile(os.path.abspath(self.pdf_path))
            except Exception:
                messagebox.showerror("预览失败", "无法打开 PDF 文件。")

    # ----------------------------------------------------------
    # 页面范围计算
    # ----------------------------------------------------------
    def _compute_page_indices(self):
        """根据页面范围设置计算要预览的 0-based 页码列表"""
        if parse_page_range is None:
            # 无解析函数时默认全部
            return list(range(self.total_pages))

        if self.page_range == "custom":
            range_str = self.custom_range_str.strip() or "all"
        else:
            range_str = self.page_range

        try:
            indices = parse_page_range(range_str, self.total_pages)
            return indices if indices else []
        except Exception:
            return list(range(self.total_pages))

    # ----------------------------------------------------------
    # UI 构建
    # ----------------------------------------------------------
    def _build_ui(self):
        self.top = tk.Toplevel(self.parent)
        self.top.title(f"打印预览 - {os.path.basename(self.pdf_path)}")
        # 适配屏幕，但限制最大尺寸
        screen_w = self.top.winfo_screenwidth()
        screen_h = self.top.winfo_screenheight()
        w = min(900, int(screen_w * 0.7))
        h = min(720, int(screen_h * 0.8))
        self.top.geometry(f"{w}x{h}")
        self.top.minsize(600, 450)
        self.top.transient(self.parent)
        self.top.grab_set()

        # 顶部工具栏
        toolbar = ttk.Frame(self.top, padding=(8, 6))
        toolbar.grid(row=0, column=0, sticky="ew")

        self.prev_btn = ttk.Button(toolbar, text="◀ 上一页", width=10,
                                   command=self._go_prev)
        self.prev_btn.pack(side="left", padx=(0, 4))

        self.next_btn = ttk.Button(toolbar, text="下一页 ▶", width=10,
                                   command=self._go_next)
        self.next_btn.pack(side="left", padx=(0, 12))

        # 跳页
        ttk.Label(toolbar, text="跳到:").pack(side="left")
        self.goto_var = tk.StringVar()
        goto_entry = ttk.Entry(toolbar, textvariable=self.goto_var, width=6)
        goto_entry.pack(side="left", padx=(2, 2))
        goto_entry.bind("<Return>", self._on_goto)
        ttk.Button(toolbar, text="Go", width=4, command=self._on_goto).pack(side="left", padx=(0, 12))

        # 缩放
        ttk.Button(toolbar, text="缩放 -", width=7,
                   command=self._zoom_out).pack(side="left", padx=(0, 2))
        self.zoom_label = ttk.Label(toolbar, text="100%", width=6, anchor="center")
        self.zoom_label.pack(side="left", padx=(2, 2))
        ttk.Button(toolbar, text="缩放 +", width=7,
                   command=self._zoom_in).pack(side="left", padx=(0, 2))
        ttk.Button(toolbar, text="适合宽度", width=9,
                   command=self._fit_width).pack(side="left", padx=(0, 12))

        # 装订标记开关
        self.bind_var = tk.BooleanVar(value=self.show_bind_marker)
        bind_cb = ttk.Checkbutton(toolbar, text="装订边标记",
                                  variable=self.bind_var,
                                  command=self._toggle_bind)
        bind_cb.pack(side="left", padx=(0, 12))
        if self.duplex_mode == "none":
            bind_cb.state(["disabled"])

        # 右侧关闭
        ttk.Button(toolbar, text="关闭", width=8,
                   command=self.top.destroy).pack(side="right")

        # 中间可滚动区域
        mid = ttk.Frame(self.top)
        mid.grid(row=1, column=0, sticky="nsew")
        self.top.columnconfigure(0, weight=1)
        self.top.rowconfigure(1, weight=1)
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(mid, bg="#3a3a3a", highlightthickness=0)
        vbar = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        hbar = ttk.Scrollbar(mid, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")

        # 鼠标滚轮支持
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", lambda e: self._render_current())

        # 底部状态栏
        self.status_var = tk.StringVar(value="正在加载...")
        status_bar = ttk.Frame(self.top, padding=(8, 4))
        status_bar.grid(row=2, column=0, sticky="ew")
        ttk.Label(status_bar, textvariable=self.status_var,
                  anchor="w").pack(side="left", fill="x", expand=True)

        # 双面模式提示标签
        mode_text = self._duplex_label()
        if mode_text:
            ttk.Label(status_bar, text=mode_text,
                      foreground=self.BIND_COLOR).pack(side="right")

        # 窗口关闭时清理
        self.top.protocol("WM_DELETE_WINDOW", self._on_close)

        # 键盘快捷键
        self.top.bind("<Left>", lambda e: self._go_prev())
        self.top.bind("<Right>", lambda e: self._go_next())
        self.top.bind("<Prior>", lambda e: self._go_prev())   # PageUp
        self.top.bind("<Next>", lambda e: self._go_next())    # PageDown

    def _duplex_label(self):
        """返回双面模式的中文标签"""
        return {
            "long": "长边翻转（书本式）",
            "short": "短边翻转（记事本式）",
            "none": "单面打印",
        }.get(self.duplex_mode, "")

    # ----------------------------------------------------------
    # 翻页
    # ----------------------------------------------------------
    def _go_prev(self):
        if self.current_pos > 0:
            self.current_pos -= 1
            self._render_current()

    def _go_next(self):
        if self.current_pos < len(self.page_indices) - 1:
            self.current_pos += 1
            self._render_current()

    def _on_goto(self, event=None):
        try:
            target = int(self.goto_var.get())
            # target 是范围内的序号（1-based）
            new_pos = target - 1
            if 0 <= new_pos < len(self.page_indices):
                self.current_pos = new_pos
                self._render_current()
            else:
                messagebox.showinfo("提示",
                    f"页码超出范围（1 - {len(self.page_indices)}）")
        except ValueError:
            messagebox.showinfo("提示", "请输入有效数字")

    def _update_nav_buttons(self):
        self.prev_btn.config(state="normal" if self.current_pos > 0 else "disabled")
        self.next_btn.config(state=(
            "normal" if self.current_pos < len(self.page_indices) - 1 else "disabled"))

    # ----------------------------------------------------------
    # 缩放
    # ----------------------------------------------------------
    def _zoom_in(self):
        if self.zoom < 3.0:
            self.zoom = min(3.0, round(self.zoom + 0.25, 2))
            self._render_current()

    def _zoom_out(self):
        if self.zoom > 0.5:
            self.zoom = max(0.5, round(self.zoom - 0.25, 2))
            self._render_current()

    def _fit_width(self):
        """根据 Canvas 宽度计算合适的 zoom"""
        canvas_w = self.canvas.winfo_width()
        if canvas_w <= 1:
            return
        page_index = self.page_indices[self.current_pos]
        try:
            doc = fitz.open(self.pdf_path)
            try:
                page = doc[page_index]
                page_w = page.rect.width
            finally:
                doc.close()
        except Exception:
            return
        if page_w <= 0:
            return
        # 目标像素宽度 = canvas_w - 2*MARGIN，base dpi=150
        # 150dpi 下 page 像素宽 = page_w * 150 / 72
        base_w = page_w * 150.0 / 72.0
        target_w = max(200, canvas_w - 2 * self.MARGIN)
        self.zoom = max(0.5, min(3.0, round(target_w / base_w, 2)))
        self._render_current()

    def _toggle_bind(self):
        self.show_bind_marker = self.bind_var.get()
        self._draw_overlay()

    # ----------------------------------------------------------
    # 渲染（后台线程）
    # ----------------------------------------------------------
    def _render_current(self):
        """触发当前页的后台渲染"""
        if not hasattr(self, "canvas"):
            return
        with self._render_lock:
            self._render_token += 1
            token = self._render_token

        page_index = self.page_indices[self.current_pos]
        self.status_var.set("渲染中...")
        self.zoom_label.config(text=f"{int(self.zoom * 100)}%")

        def worker():
            png_data = render_pdf_page(self.pdf_path, page_index,
                                       dpi=150, zoom=self.zoom)
            # 回到主线程
            try:
                self.top.after(0, lambda: self._on_render_done(token, png_data))
            except Exception:
                pass  # 窗口已关闭

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _on_render_done(self, token, png_data):
        """渲染完成回调（主线程）"""
        # 丢弃过期请求
        if token != self._render_token:
            return
        if png_data is None:
            self.status_var.set("渲染失败")
            self.canvas.delete("all")
            self.canvas.create_text(
                self.canvas.winfo_width() // 2,
                self.canvas.winfo_height() // 2,
                text="页面渲染失败", fill="white",
                font=("Segoe UI", 12))
            self._update_nav_buttons()
            return

        # 创建 PhotoImage
        self._current_photo = tk.PhotoImage(data=png_data)
        img_w = self._current_photo.width()
        img_h = self._current_photo.height()

        # 设置 scrollregion
        scroll_w = img_w + 2 * self.MARGIN
        scroll_h = img_h + 2 * self.MARGIN
        self.canvas.configure(scrollregion=(0, 0, scroll_w, scroll_h))
        self.canvas.delete("all")

        # 居中放置图像
        x = self.MARGIN
        y = self.MARGIN
        self.canvas.create_image(x, y, anchor="nw", image=self._current_photo)

        # 绘制装订边标记
        self._draw_overlay(img_w, img_h, x, y)

        # 更新状态栏
        orig_page = self.page_indices[self.current_pos] + 1  # 1-based 显示
        self.status_var.set(
            f"第 {self.current_pos + 1} / {len(self.page_indices)} 页"
            f"   |   原始页码: {orig_page} / {self.total_pages}"
            f"   |   缩放: {int(self.zoom * 100)}%"
        )
        self._update_nav_buttons()

    def _draw_overlay(self, img_w=None, img_h=None, x=None, y=None):
        """绘制双面装订边标记"""
        if not self.show_bind_marker:
            return
        if self.duplex_mode not in ("long", "short"):
            return
        if img_w is None:
            # 仅重绘标记时（开关切换），无法获取尺寸则跳过
            return

        dash = (6, 4)
        if self.duplex_mode == "long":
            # 长边翻转：左侧装订边（竖线）
            line_x = x - 10
            self.canvas.create_line(line_x, y, line_x, y + img_h,
                                    fill=self.BIND_COLOR, width=2,
                                    dash=dash)
            self.canvas.create_text(line_x, y - 8, text="装订边",
                                    fill=self.BIND_COLOR,
                                    font=("Segoe UI", 9), anchor="s")
        elif self.duplex_mode == "short":
            # 短边翻转：上侧装订边（横线）
            line_y = y - 10
            self.canvas.create_line(x, line_y, x + img_w, line_y,
                                    fill=self.BIND_COLOR, width=2,
                                    dash=dash)
            self.canvas.create_text(x - 8, line_y, text="装订边",
                                    fill=self.BIND_COLOR,
                                    font=("Segoe UI", 9), anchor="e")

    # ----------------------------------------------------------
    # 事件处理
    # ----------------------------------------------------------
    def _on_mousewheel(self, event):
        """鼠标滚轮垂直滚动"""
        # Windows: event.delta 是 120 的倍数
        delta = -1 * (event.delta // 120)
        self.canvas.yview_scroll(delta, "units")

    def _on_close(self):
        with self._render_lock:
            self._render_token += 1  # 使正在进行的渲染失效
        self.top.destroy()
