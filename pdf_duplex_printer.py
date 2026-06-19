"""
PDF双面打印延迟控制脚本 - 核心模块 (Windows版)
支持自动页面缩放、横竖向调整和双面打印
注：先安装acrobat，以使用acrobat的自动处理页面翻转、缩放至纸张尺寸功能

依赖库
pip install pywin32 pypdf
"""

import os
import sys
import time
import shutil
import tempfile
import subprocess
import threading
import winreg
from typing import Callable, Optional

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("请先安装pypdf库: pip install pypdf")
    sys.exit(1)

try:
    import win32print
    import win32api
    import win32con
except ImportError:
    print("请先安装pywin32库: pip install pywin32")
    sys.exit(1)


# ============================================================
# 常量定义
# ============================================================

# 双面打印模式映射 (DEVMODE Duplex 值)
DUPLEX_MODES = {
    "long": 2,    # DMDUP_VERTICAL - 长边翻转（书本式）
    "short": 3,   # DMDUP_HORIZONTAL - 短边翻转（记事本式）
    "none": 1,    # DMDUP_SIMPLEX - 单面
}

# 双面模式可读名称（反向映射）
DUPLEX_NAMES = {v: k for k, v in {
    "单面": 1,
    "双面(长边)": 2,
    "双面(短边)": 3,
}.items()}

# 纸张大小代码 -> 可读名称映射表 (Windows DMPAPER 常量)
PAPER_SIZE_MAP = {
    1: "Letter", 2: "Letter Small", 3: "Tabloid", 4: "Ledger",
    5: "Legal", 6: "Statement", 7: "Executive", 8: "A3",
    9: "A4", 10: "A4 Small", 11: "A5", 12: "B4 (JIS)",
    13: "B5 (JIS)", 14: "Folio", 15: "Quarto", 16: "10x14",
    17: "11x17", 18: "Note", 19: "Envelope #9", 20: "Envelope #10",
    21: "Envelope #11", 22: "Envelope #12", 23: "Envelope #14",
    24: "C size sheet", 25: "D size sheet", 26: "E size sheet",
    27: "Envelope DL", 28: "Envelope C5", 29: "Envelope C3",
    30: "Envelope C4", 31: "Envelope C6", 32: "Envelope C65",
    33: "Envelope B4", 34: "Envelope B5", 35: "Envelope B6",
    36: "Envelope Italy", 37: "Envelope Monarch", 38: "Envelope Personal",
    39: "Fanfold US Std", 40: "Fanfold German Std", 41: "Fanfold German Legal",
    66: "Envelope Invite", 67: "A2", 68: "A6", 69: "B6 (JIS)",
    70: "B5 (ISO)", 71: "A1", 72: "A0", 73: "8K", 74: "10K",
}

# 回调事件类型:
#   "info"          - 一般信息 {message}
#   "split_start"   - 开始拆分 {total_pages, sheets}
#   "split_done"    - 拆分完成 {sheets, temp_dir}
#   "config"        - 打印机配置 {printer, duplex_mode, success}
#   "batch_start"   - 批次开始 {index, total, filename}
#   "batch_done"    - 批次完成 {index, total}
#   "batch_fail"    - 批次失败 {index, total, error}
#   "delay"         - 延迟倒计时 {remaining, total}
#   "cleanup"       - 清理临时文件 {temp_dir, kept}
#   "done"          - 全部完成 {success_count, total_sheets, cancelled}
#   "error"         - 错误 {message}
ProgressCallback = Optional[Callable[[str, dict], None]]


# ============================================================
# 打印机相关函数
# ============================================================

def get_available_printers() -> list:
    """获取所有可用的打印机名称列表"""
    printers = []
    for printer in win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    ):
        printers.append(printer[2])
    return printers


def get_default_printer() -> str:
    """获取系统默认打印机名称"""
    return win32print.GetDefaultPrinter()


def get_paper_size_name(paper_size_code: int) -> str:
    """将纸张大小代码映射为可读名称"""
    return PAPER_SIZE_MAP.get(paper_size_code, f"Unknown ({paper_size_code})")


def get_duplex_name(duplex_code: int) -> str:
    """将双面打印代码映射为可读名称"""
    names = {1: "单面", 2: "双面(长边)", 3: "双面(短边)"}
    return names.get(duplex_code, f"Unknown ({duplex_code})")


def resolve_printer(printer_input: str) -> tuple:
    """
    解析打印机名称，支持三种匹配模式:
    1. 精确匹配 - 输入完全等于某个打印机名
    2. 数字编号 - 输入为纯数字，作为打印机列表的索引（1-based）
    3. 模糊匹配 - 大小写不敏感的子串匹配

    返回: (matched_printer: str|None, matches: list)
      - 精确/编号/单匹配: (printer_name, [printer_name])
      - 多匹配: (None, [match1, match2, ...])
      - 无匹配: (None, [])
    """
    printers = get_available_printers()

    # 1. 精确匹配
    if printer_input in printers:
        return printer_input, [printer_input]

    # 2. 数字编号（1-based）
    if printer_input.isdigit():
        idx = int(printer_input)
        if 1 <= idx <= len(printers):
            return printers[idx - 1], [printers[idx - 1]]
        return None, []

    # 3. 模糊匹配（大小写不敏感子串匹配）
    lower_input = printer_input.lower()
    matches = [p for p in printers if lower_input in p.lower()]

    if len(matches) == 1:
        return matches[0], matches
    elif len(matches) > 1:
        return None, matches
    else:
        return None, []


def interactive_select_printer() -> Optional[str]:
    """
    交互式选择打印机（CLI 用）
    显示编号列表，提示用户输入序号或名称片段
    支持回车选择默认打印机

    返回: 打印机名称 或 None（用户取消）
    """
    printers = get_available_printers()
    default_printer = get_default_printer()

    print("\n可用的打印机:")
    print("-" * 50)
    for i, printer in enumerate(printers, 1):
        marker = " (默认)" if printer == default_printer else ""
        print(f"  {i}. {printer}{marker}")
    print("-" * 50)

    while True:
        try:
            user_input = input(
                "\n请输入打印机序号或名称片段（回车=默认打印机, q=取消）: "
            ).strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return None

        if user_input.lower() == 'q':
            return None

        if not user_input:
            return default_printer

        matched, matches = resolve_printer(user_input)

        if matched:
            return matched

        if matches:
            # 多匹配，进一步选择
            print(f"\n找到多个匹配的打印机:")
            for i, p in enumerate(matches, 1):
                marker = " (默认)" if p == default_printer else ""
                print(f"  {i}. {p}{marker}")
            try:
                choice = input("请输入序号选择（回车=取消）: ").strip()
                if not choice:
                    continue
                idx = int(choice)
                if 1 <= idx <= len(matches):
                    return matches[idx - 1]
                print("序号超出范围，请重新输入。")
            except (ValueError, EOFError, KeyboardInterrupt):
                print("输入无效，请重新输入。")
        else:
            print(f"未找到匹配的打印机，请重新输入。")


def show_printer_capabilities(printer_name: str) -> dict:
    """
    获取打印机当前配置信息

    返回: {
        "name": str,
        "duplex": str,       # 可读名称
        "paper_size": str,   # 可读名称
        "orientation": str,  # "纵向" | "横向"
    }
    """
    result = {
        "name": printer_name,
        "duplex": "未知",
        "paper_size": "未知",
        "orientation": "未知",
    }

    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            properties = win32print.GetPrinter(handle, 2)
            pDevMode = properties["pDevMode"]

            if pDevMode:
                result["duplex"] = get_duplex_name(pDevMode.Duplex)
                result["paper_size"] = get_paper_size_name(pDevMode.PaperSize)
                result["orientation"] = "横向" if pDevMode.Orientation == 2 else "纵向"
        finally:
            win32print.ClosePrinter(handle)
    except Exception as e:
        result["error"] = str(e)

    return result


def configure_printer_duplex(printer_name: str, duplex_mode: str = "long") -> bool:
    """
    配置打印机双面打印设置
    注意: 这个函数修改打印机的默认设置

    参数:
        printer_name: 打印机名称
        duplex_mode: "long"=长边翻转, "short"=短边翻转, "none"=单面
    """
    try:
        handle = win32print.OpenPrinter(printer_name)

        try:
            properties = win32print.GetPrinter(handle, 2)
            pDevMode = properties["pDevMode"]

            if pDevMode:
                # 设置双面打印
                pDevMode.Duplex = DUPLEX_MODES.get(duplex_mode, 2)
                # 设置纸张大小为A4
                pDevMode.PaperSize = 9  # DMPAPER_A4

                # 应用设置
                properties["pDevMode"] = pDevMode
                win32print.SetPrinter(handle, 2, properties, 0)
                return True
            else:
                return False

        finally:
            win32print.ClosePrinter(handle)

    except Exception:
        return False


# ============================================================
# Acrobat Reader 查找
# ============================================================

def find_acrobat_reader() -> Optional[str]:
    """
    查找 Adobe Acrobat Reader 可执行文件路径
    依次尝试：常见安装路径 -> Windows 注册表
    """
    # 常见安装路径
    acrobat_paths = [
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
    ]

    for path in acrobat_paths:
        if os.path.exists(path):
            return path

    # 从注册表查找
    try:
        for reg_path in [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
        ]:
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE, reg_path
                )
                exe_path = winreg.QueryValue(key, None)
                winreg.CloseKey(key)
                if exe_path and os.path.exists(exe_path):
                    return exe_path
            except (FileNotFoundError, OSError):
                continue
    except Exception:
        pass

    return None


# ============================================================
# PDF 拆分
# ============================================================

def split_pdf_for_duplex(
    input_pdf: str,
    output_dir: str,
    progress_callback: ProgressCallback = None,
) -> list:
    """
    将PDF按双面打印需求分割成多个小文件
    每个文件包含2页（一张纸的正反面）

    参数:
        input_pdf: 输入PDF文件路径
        output_dir: 输出目录（由调用方创建）
        progress_callback: 进度回调函数

    返回: 临时文件路径列表
    """
    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)
    sheets = (total_pages + 1) // 2

    if progress_callback:
        progress_callback("split_start", {
            "total_pages": total_pages,
            "sheets": sheets,
        })

    temp_files = []
    for i in range(0, total_pages, 2):
        writer = PdfWriter()

        # 添加正面（奇数页）
        writer.add_page(reader.pages[i])

        # 添加反面（偶数页），如果存在
        if i + 1 < total_pages:
            writer.add_page(reader.pages[i + 1])

        temp_file = os.path.join(output_dir, f"page_{i // 2 + 1:03d}.pdf")
        with open(temp_file, "wb") as f:
            writer.write(f)

        temp_files.append(temp_file)

    if progress_callback:
        progress_callback("split_done", {
            "sheets": sheets,
            "temp_dir": output_dir,
        })

    return temp_files


# ============================================================
# 打印
# ============================================================

def print_pdf_advanced(pdf_file: str, printer_name: str = None) -> bool:
    """
    使用Acrobat Reader进行高级打印
    支持自动缩放、双面打印等功能

    参数:
        pdf_file: PDF文件路径
        printer_name: 打印机名称，None表示使用默认打印机

    返回: 是否成功发送打印命令
    """
    try:
        if printer_name is None:
            printer_name = get_default_printer()

        abs_path = os.path.abspath(pdf_file)
        acrobat_path = find_acrobat_reader()

        if acrobat_path:
            try:
                subprocess.Popen(
                    [acrobat_path, "/t", abs_path, printer_name],
                    shell=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                time.sleep(3)
            except Exception:
                # 降级到ShellExecute方法
                win32api.ShellExecute(
                    0, "print", abs_path, f'/d:"{printer_name}"', ".", 0
                )
        else:
            # 未找到Adobe Reader，使用系统默认打印
            win32api.ShellExecute(
                0, "print", abs_path, f'/d:"{printer_name}"', ".", 0
            )

        return True

    except Exception:
        return False


# ============================================================
# 主打印流程
# ============================================================

def print_with_delay(
    pdf_file: str,
    delay_seconds: int = 15,
    printer_name: Optional[str] = None,
    keep_temp: bool = False,
    configure_duplex: bool = True,
    duplex_mode: str = "long",
    progress_callback: ProgressCallback = None,
    cancel_event: Optional[threading.Event] = None,
) -> dict:
    """
    主函数：分批打印PDF，每次打印后等待指定时间

    参数:
        pdf_file: 输入PDF文件路径
        delay_seconds: 每次打印后的等待时间（秒）
        printer_name: 打印机名称（None表示使用默认打印机）
        keep_temp: 是否保留临时文件
        configure_duplex: 是否自动配置双面打印
        duplex_mode: "long"=长边翻转, "short"=短边翻转, "none"=单面
        progress_callback: 进度回调函数
        cancel_event: 取消事件（ threading.Event）

    返回: {"success_count": int, "total_sheets": int, "cancelled": bool}
    """
    def _cb(event_type: str, data: dict = None):
        if progress_callback:
            progress_callback(event_type, data or {})

    def _is_cancelled() -> bool:
        return cancel_event is not None and cancel_event.is_set()

    result = {"success_count": 0, "total_sheets": 0, "cancelled": False}

    if not os.path.exists(pdf_file):
        _cb("error", {"message": f"文件不存在: {pdf_file}"})
        return result

    # 确定打印机
    if printer_name is None:
        printer_name = get_default_printer()
        _cb("info", {"message": f"使用默认打印机: {printer_name}"})
    else:
        _cb("info", {"message": f"使用打印机: {printer_name}"})

    # 配置双面打印
    if configure_duplex:
        _cb("info", {"message": "配置打印机双面打印设置..."})
        success = configure_printer_duplex(printer_name, duplex_mode)
        _cb("config", {
            "printer": printer_name,
            "duplex_mode": duplex_mode,
            "success": success,
        })

    # 使用系统临时目录创建临时子目录
    temp_dir = tempfile.mkdtemp(prefix="pdf_duplex_")
    _cb("info", {"message": f"临时目录: {temp_dir}"})

    try:
        # 检查取消
        if _is_cancelled():
            result["cancelled"] = True
            return result

        # 分割PDF
        temp_files = split_pdf_for_duplex(pdf_file, temp_dir, progress_callback)
        total_sheets = len(temp_files)
        result["total_sheets"] = total_sheets

        if total_sheets == 0:
            _cb("error", {"message": "PDF没有页面可打印"})
            return result

        _cb("info", {"message": f"开始打印，共 {total_sheets} 批次..."})

        # 逐批打印
        for idx, temp_file in enumerate(temp_files, 1):
            # 检查取消
            if _is_cancelled():
                result["cancelled"] = True
                _cb("info", {"message": "用户已取消打印"})
                break

            _cb("batch_start", {
                "index": idx,
                "total": total_sheets,
                "filename": os.path.basename(temp_file),
            })

            success = print_pdf_advanced(temp_file, printer_name)

            if success:
                result["success_count"] += 1
                _cb("batch_done", {
                    "index": idx,
                    "total": total_sheets,
                })
            else:
                _cb("batch_fail", {
                    "index": idx,
                    "total": total_sheets,
                    "error": "打印命令发送失败",
                })

            # 延迟等待（逐秒倒计时，支持取消）
            if idx < total_sheets or True:  # 每批都等待，确保Reader读取文件
                for remaining in range(delay_seconds, 0, -1):
                    if _is_cancelled():
                        result["cancelled"] = True
                        break
                    _cb("delay", {
                        "remaining": remaining,
                        "total": delay_seconds,
                    })
                    time.sleep(1)

                if _is_cancelled():
                    _cb("info", {"message": "用户已取消打印"})
                    break

        # 额外等待确保最后一个文件被Reader完全读取
        if not _is_cancelled():
            _cb("info", {"message": "等待最后的打印作业完成..."})
            time.sleep(5)

    finally:
        # 清理临时文件
        if not keep_temp:
            _cb("cleanup", {"temp_dir": temp_dir, "kept": False})
            shutil.rmtree(temp_dir, ignore_errors=True)
        else:
            _cb("cleanup", {"temp_dir": temp_dir, "kept": True})
            _cb("info", {"message": f"临时文件保存在: {temp_dir}"})

    _cb("done", {
        "success_count": result["success_count"],
        "total_sheets": result["total_sheets"],
        "cancelled": result["cancelled"],
    })

    return result


# ============================================================
# CLI 默认回调
# ============================================================

def default_cli_callback(event_type: str, data: dict) -> None:
    """CLI 默认进度回调实现，将事件转为 print 输出"""

    if event_type == "info":
        print(data.get("message", ""))

    elif event_type == "split_start":
        print(f"PDF总页数: {data['total_pages']}")
        print(f"需要打印 {data['sheets']} 张纸（双面）")

    elif event_type == "split_done":
        print(f"PDF拆分完成，共 {data['sheets']} 个临时文件")

    elif event_type == "config":
        mode_names = {"long": "长边翻转", "short": "短边翻转", "none": "单面"}
        mode_name = mode_names.get(data["duplex_mode"], data["duplex_mode"])
        if data["success"]:
            print(f"  已配置打印机为 {mode_name} 模式")
        else:
            print(f"  警告: 打印机配置失败，将使用当前设置")

    elif event_type == "batch_start":
        print(f"\n[{data['index']}/{data['total']}] 正在打印: {data['filename']}")

    elif event_type == "batch_done":
        print(f"  打印命令已发送")

    elif event_type == "batch_fail":
        print(f"  警告: 打印失败 - {data.get('error', '未知错误')}")

    elif event_type == "delay":
        remaining = data["remaining"]
        total = data["total"]
        # 使用 \r 实现同一行刷新倒计时
        sys.stdout.write(f"\r  等待 {total} 秒... 剩余 {remaining} 秒  ")
        sys.stdout.flush()
        if remaining == 1:
            print()  # 倒计时结束换行

    elif event_type == "cleanup":
        if not data["kept"]:
            print("清理临时文件...")

    elif event_type == "done":
        success = data["success_count"]
        total = data["total_sheets"]
        cancelled = data["cancelled"]
        if cancelled:
            print(f"\n打印已取消! 已完成 {success}/{total} 批次")
        else:
            print(f"\n打印任务完成! 成功发送 {success}/{total} 批次")

    elif event_type == "error":
        print(f"错误: {data.get('message', '未知错误')}")
