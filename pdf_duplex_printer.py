"""
PDF双面打印延迟控制脚本 - 核心模块 (Windows版)
支持自动页面缩放、横竖向调整和双面打印

打印引擎优先级 (auto模式): SumatraPDF > Adobe Acrobat > 系统默认(ShellExecute)
1. SumatraPDF（推荐）:
   - 同步打印，精确控制缩放和双面模式
   - 通过 -print-settings 控制双面(fit)和缩放(duplexlong/short/simplex)
   - 不需要 DEVMODE 配置，不修改打印机全局设置
   - 有退出码用于错误判断
2. Adobe Acrobat Reader:
   - 异步打印（Popen + sleep），双面通过 DEVMODE 配置
   - 需要安装 Adobe Acrobat Reader
3. 系统默认(ShellExecute):
   - 使用系统默认关联程序打印，降级方案

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

# Fix Windows console encoding for Chinese characters
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

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

# 打印引擎类型常量
ENGINE_SUMATRA = "sumatra"
ENGINE_ACROBAT = "acrobat"
ENGINE_SHELL = "shell"

# 引擎可读名称
ENGINE_NAMES = {
    "sumatra": "SumatraPDF",
    "acrobat": "Acrobat Reader",
    "shell": "系统默认(ShellExecute)",
}

# SumatraPDF 双面模式 -> -print-settings 参数映射
SUMATRA_DUPLEX_MAP = {
    "long": "duplexlong",
    "short": "duplexshort",
    "none": "simplex",
}

# 回调事件类型:
#   "info"          - 一般信息 {message}
#   "engine"        - 引擎选择结果 {engine, engine_name, engine_path}
#   "copy_start"    - 开始打印新一份 {copy_index, total_copies}
#   "split_start"   - 开始拆分 {total_pages, sheets} (is_duplex: bool)
#   "split_done"    - 拆分完成 {sheets, temp_dir}
#   "config"        - 打印机配置 {printer, duplex_mode, success}
#   "batch_start"   - 批次开始 {index, total, filename}
#   "batch_done"    - 批次完成 {index, total}
#   "batch_fail"    - 批次失败 {index, total, error}
#   "delay"         - 延迟倒计时 {remaining, total}
#   "file_start"    - 开始处理某个文件（多文件模式）{index, total, filepath}
#   "file_done"     - 某个文件处理完成 {index, total, success_count, total_sheets}
#   "cleanup"       - 清理临时文件 {temp_dir, kept}
#   "done"          - 全部完成 {success_count, total_sheets, cancelled}
#   "stats"         - 打印统计 {total_pages, elapsed_seconds, success_rate, ...}
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
        "is_network": bool,  # 是否为网络打印机
        "model_info": dict,  # 型号检测结果
    }
    """
    result = {
        "name": printer_name,
        "duplex": "未知",
        "paper_size": "未知",
        "orientation": "未知",
        "is_network": False,
        "model_info": {},
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

    # 检测网络打印机
    result["is_network"] = is_network_printer(printer_name)

    # 检测打印机型号
    result["model_info"] = detect_printer_model(printer_name)

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
# 打印机型号数据库 (v0.4)
# ============================================================

def _load_printer_models() -> dict:
    """加载打印机型号数据库"""
    import json
    model_paths = [
        os.path.join(_get_app_dir(), "printer_models.json"),
        # PyInstaller 打包后数据文件在 sys._MEIPASS (internal/ 目录)
        os.path.join(sys._MEIPASS, "printer_models.json") if getattr(sys, 'frozen', False) else "",
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "printer_models.json"),
    ]
    for path in model_paths:
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
    return {}


def _match_model(printer_name: str, models: dict) -> tuple:
    """
    在数据库中匹配打印机型号

    返回: (brand, model_key, model_info) 或 (None, None, {})
    """
    name_upper = printer_name.upper()

    for brand, brand_models in models.items():
        if brand in ("_description", "_version", "_note", "generic", "network_penalty", "keywords"):
            continue
        if not isinstance(brand_models, dict):
            continue
        for model_key, model_info in brand_models.items():
            if not isinstance(model_info, dict):
                continue
            if model_key.upper() in name_upper:
                return (brand, model_key, model_info)

    return (None, None, {})


def _guess_printer_type(printer_name: str, models: dict) -> str:
    """
    根据打印机名称猜测打印机类型

    返回: "laser_old" | "laser_new" | "inkjet" | "unknown"
    """
    name_upper = printer_name.upper()
    keywords = models.get("keywords", {})

    # 老型号关键词
    for kw in keywords.get("old_models", []):
        if kw.upper() in name_upper:
            return "laser_old"

    # 新型号关键词
    for kw in keywords.get("new_models", []):
        if kw.upper() in name_upper:
            return "laser_new"

    # 喷墨关键词
    for kw in keywords.get("inkjet_models", []):
        if kw.upper() in name_upper:
            return "inkjet"

    # 根据品牌猜测
    for brand_keyword in ["LASERJET", "LASER", "LBP"]:
        if brand_keyword in name_upper:
            return "laser_old"

    for brand_keyword in ["INKJET", "DESKJET", "OFFICEJET"]:
        if brand_keyword in name_upper:
            return "inkjet"

    return "unknown"


def detect_printer_model(printer_name: str) -> dict:
    """
    检测打印机型号并返回推荐配置

    返回: {
        "matched": bool,
        "brand": str or None,
        "model": str or None,
        "recommended_delay": int,
        "note": str,
        "verified": bool,
        "printer_type": str,  # known/laser_old/laser_new/inkjet/unknown
    }
    """
    models = _load_printer_models()
    brand, model_key, model_info = _match_model(printer_name, models)

    if model_info:
        return {
            "matched": True,
            "brand": brand,
            "model": model_key,
            "recommended_delay": model_info.get("delay", 15),
            "note": model_info.get("note", ""),
            "verified": model_info.get("verified", False),
            "printer_type": "known",
        }

    # 未匹配到精确型号，使用通用猜测
    printer_type = _guess_printer_type(printer_name, models)
    generic = models.get("generic", {}).get(printer_type, {})

    return {
        "matched": False,
        "brand": None,
        "model": None,
        "recommended_delay": generic.get("delay", 15),
        "note": generic.get("note", f"未匹配到已知型号，使用通用建议（{printer_type}）"),
        "verified": False,
        "printer_type": printer_type,
    }


def get_recommended_delay(printer_name: str, user_delay: int = None) -> int:
    """
    获取推荐延迟时间

    参数:
        printer_name: 打印机名称
        user_delay: 用户手动指定的延迟（None表示未指定）

    返回: 推荐的延迟秒数
          - 如果用户指定了延迟，优先使用用户设置
          - 否则使用数据库推荐的延迟
    """
    if user_delay is not None:
        return user_delay

    model_info = detect_printer_model(printer_name)
    base_delay = model_info.get("recommended_delay", 15)

    # 检测是否为网络打印机
    if is_network_printer(printer_name):
        models = _load_printer_models()
        network_penalty = models.get("network_penalty", {}).get("delay", 5)
        base_delay += network_penalty

    return base_delay


# ============================================================
# 网络打印机检测 (v0.4)
# ============================================================

def is_network_printer(printer_name: str) -> bool:
    """
    检测打印机是否为网络打印机

    通过检查打印机端口类型判断：
    - IP端口（如 192.168.x.x）-> 网络打印机
    - WSD端口 -> 网络打印机
    - TCP/IP端口 -> 网络打印机
    - USB端口 -> 本地打印机
    - LPT端口 -> 本地打印机

    返回: True 表示可能是网络打印机
    """
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            properties = win32print.GetPrinter(handle, 2)
            port_name = properties.get("pPortName", "").upper() if properties else ""

            # 网络打印机端口特征
            network_indicators = [
                "IP_",           # 标准TCP/IP端口
                "WSD",           # Web Services for Devices
                "HTTP:",         # HTTP打印
                "HTTPS:",        # HTTPS打印
                "IPP",           # Internet Printing Protocol
                "LPR",           # Line Printer Remote
            ]

            # IP地址模式
            import re
            if re.match(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', port_name):
                return True

            for indicator in network_indicators:
                if port_name.startswith(indicator):
                    return True

            return False

        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        return False


# ============================================================
# 打印引擎查找
# ============================================================

def _get_app_dir() -> str:
    """
    获取应用程序目录（兼容开发环境和 PyInstaller 打包环境）
    开发环境: 返回脚本所在目录
    打包环境: 返回 exe 所在目录
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


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


def find_sumatra_pdf() -> Optional[str]:
    """
    查找 SumatraPDF 可执行文件路径
    依次尝试：vendor目录(打包内置) -> 系统PATH -> 常见安装路径 -> 注册表
    返回: 找到返回完整路径，未找到返回 None
    """
    # 1. 检查打包在 vendor 目录中的 SumatraPDF.exe
    app_dir = _get_app_dir()
    vendor_path = os.path.join(app_dir, "vendor", "SumatraPDF.exe")
    if os.path.exists(vendor_path):
        return vendor_path

    # 2. 检查系统 PATH
    for exe_name in ("sumatrapdf.exe", "SumatraPDF.exe"):
        exe_path = shutil.which(exe_name)
        if exe_path:
            return exe_path

    # 3. 检查常见安装路径
    for path in (
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
    ):
        if os.path.exists(path):
            return path

    # 4. 从注册表查找
    for reg_path in (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\SumatraPDF.exe",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\SumatraPDF.exe",
    ):
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            exe_path = winreg.QueryValue(key, None)
            winreg.CloseKey(key)
            if exe_path and os.path.exists(exe_path):
                return exe_path
        except (FileNotFoundError, OSError):
            continue

    return None


# ============================================================
# 打印队列管理 (v0.4)
# ============================================================

def _parse_job_status(status_value):
    """将 JOB_STATUS_* 位掩码解析为中文状态描述"""
    if status_value == 0:
        return "就绪"
    parts = []
    if status_value & win32print.JOB_STATUS_PAUSED:
        parts.append("暂停")
    if status_value & win32print.JOB_STATUS_ERROR:
        parts.append("错误")
    if status_value & win32print.JOB_STATUS_DELETING:
        parts.append("正在删除")
    if status_value & win32print.JOB_STATUS_SPOOLING:
        parts.append("假脱机")
    if status_value & win32print.JOB_STATUS_PRINTING:
        parts.append("打印中")
    if status_value & win32print.JOB_STATUS_OFFLINE:
        parts.append("离线")
    if status_value & win32print.JOB_STATUS_PAPEROUT:
        parts.append("缺纸")
    if status_value & win32print.JOB_STATUS_USER_INTERVENTION:
        parts.append("用户干预")
    if status_value & win32print.JOB_STATUS_BLOCKED_DEVQ:
        parts.append("队列阻塞")
    if status_value & win32print.JOB_STATUS_DELETED:
        parts.append("已删除")
    if status_value & win32print.JOB_STATUS_PRINTED:
        parts.append("已打印")
    if status_value & win32print.JOB_STATUS_RESTART:
        parts.append("需重启")
    if status_value & win32print.JOB_STATUS_COMPLETE:
        parts.append("完成")
    return ", ".join(parts) if parts else f"0x{status_value:08X}"


def _format_submitted(submitted_time):
    """格式化提交时间"""
    try:
        if submitted_time:
            return submitted_time.strftime("%H:%M:%S")
    except Exception:
        pass
    return ""


def _normalize_job(job_info, level):
    """
    将 JOB_INFO_1 或 JOB_INFO_2 统一为标准 dict。
    level 1/2 主要字段名一致（pDocument/TotalPages/pUserName/Status/JobId/Size/Submitted），
    但 level=2 额外含 pPrinterName/pMachineName/pDriverName 等字段。
    """
    status_value = job_info.get("Status", 0)
    return {
        "job_id": job_info.get("JobId", 0),
        "document": job_info.get("pDocument", "未知"),
        "status": _parse_job_status(status_value),
        "pages": job_info.get("TotalPages", 0),
        "size": job_info.get("Size", 0),
        "submitted": _format_submitted(job_info.get("Submitted", None)),
        "owner": job_info.get("pUserName", ""),
    }


def _list_jobs_via_enumjobs(printer_name, level):
    """
    用 win32print.EnumJobs 指定 level 获取作业。
    返回 (jobs_list, error_str)；成功时 error_str 为空。
    """
    jobs = []
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            job_infos = win32print.EnumJobs(handle, 0, -1, level)
            for ji in job_infos:
                jobs.append(_normalize_job(ji, level))
        finally:
            win32print.ClosePrinter(handle)
        return jobs, ""
    except Exception as e:
        return [], str(e)


def _list_jobs_via_wmi(printer_name):
    """
    WMI 降级方案：通过 win32com 查询 Win32_PrintJob。
    返回 (jobs_list, error_str)。
    无需额外依赖（pywin32 自带 win32com）。
    """
    jobs = []
    try:
        import win32com.client
        wmi = win32com.client.GetObject("winmgmts:")
        # Win32_PrintJob.Name 字段格式为 "打印机名, 作业ID"
        query = "SELECT * FROM Win32_PrintJob"
        for pj in wmi.ExecQuery(query):
            name = str(pj.Name or "")
            # 按打印机名过滤（Name = "PrinterName, JobId"）
            if not name.startswith(printer_name + ","):
                continue
            # 提取 job_id
            try:
                job_id = int(name.split(",", 1)[1])
            except Exception:
                job_id = 0
            # 状态码映射（Win32_PrintJob.JobStatus 是字符串，StatusCode 是数字）
            status_str = str(pj.JobStatus or "") if hasattr(pj, "JobStatus") else ""
            if not status_str:
                sc = int(pj.StatusCode or 0) if hasattr(pj, "StatusCode") else 0
                status_str = _wmi_status_to_text(sc)
            jobs.append({
                "job_id": job_id,
                "document": str(pj.Document or "未知"),
                "status": status_str or "就绪",
                "pages": int(pj.TotalPages or 0) if hasattr(pj, "TotalPages") else 0,
                "size": int(pj.Size or 0) if hasattr(pj, "Size") else 0,
                "submitted": "",
                "owner": str(pj.Owner or "") if hasattr(pj, "Owner") else "",
            })
        return jobs, ""
    except Exception as e:
        return [], str(e)


def _wmi_status_to_text(status_code):
    """将 Win32_PrintJob.StatusCode 映射为中文（简化版）"""
    mapping = {
        1: "暂停", 2: "错误", 3: "正在删除", 4: "假脱机",
        5: "打印中", 6: "离线", 7: "缺纸", 8: "已打印",
        9: "已删除", 10: "需重启",
    }
    return mapping.get(status_code, f"状态码:{status_code}")


def list_print_jobs(printer_name: str = None) -> dict:
    """
    列出指定打印机的打印队列（Windows 11 25H2 兼容版）。

    采用多级降级策略确保 25H2 下能获取到作业：
      方法1: win32print.EnumJobs level=1（传统本地打印机）
      方法2: win32print.EnumJobs level=2（25H2 IPP/WPS 队列）
      方法3: WMI 查询 Win32_PrintJob（兜底）
    任一方法返回非空即采用；全部为空时返回空列表+诊断信息。

    参数:
        printer_name: 打印机名称，None 表示使用默认打印机

    返回: {
        "jobs": [                    # 作业列表（可能为空）
            {
                "job_id": int,
                "document": str,
                "status": str,
                "pages": int,
                "size": int,
                "submitted": str,
                "owner": str,
            }, ...
        ],
        "method": str,               # 使用的获取方法: "EnumJobs-L1"/"EnumJobs-L2"/"WMI"/"none"
        "diagnostics": str,          # 诊断信息（尝试了哪些方法、各返回多少）
        "error": str,                # 最后一次异常文本（无异常为空字符串）
    }
    """
    if printer_name is None:
        printer_name = get_default_printer()

    result = {
        "jobs": [],
        "method": "none",
        "diagnostics": "",
        "error": "",
    }

    # ---- 方法1: EnumJobs level=1 ----
    jobs1, err1 = _list_jobs_via_enumjobs(printer_name, 1)
    diag_parts = [f"EnumJobs-L1({len(jobs1)}个)"]
    if jobs1:
        result["jobs"] = jobs1
        result["method"] = "EnumJobs-L1"
        result["diagnostics"] = "通过 EnumJobs level=1 获取"
        return result
    if err1:
        diag_parts[0] = f"EnumJobs-L1(失败: {err1[:40]})"
        result["error"] = err1

    # ---- 方法2: EnumJobs level=2 ----
    jobs2, err2 = _list_jobs_via_enumjobs(printer_name, 2)
    diag_parts.append(f"EnumJobs-L2({len(jobs2)}个)")
    if jobs2:
        result["jobs"] = jobs2
        result["method"] = "EnumJobs-L2"
        result["diagnostics"] = "通过 EnumJobs level=2 获取（level=1 未返回数据）"
        return result
    if err2 and not result["error"]:
        result["error"] = err2

    # ---- 方法3: WMI 降级 ----
    jobs3, err3 = _list_jobs_via_wmi(printer_name)
    diag_parts.append(f"WMI({len(jobs3)}个)")
    if jobs3:
        result["jobs"] = jobs3
        result["method"] = "WMI"
        result["diagnostics"] = "通过 WMI 查询获取（EnumJobs L1/L2 均未返回数据）"
        return result
    if err3 and not result["error"]:
        result["error"] = err3

    # ---- 全部为空 ----
    result["method"] = "none"
    if result["error"]:
        result["diagnostics"] = (
            f"尝试 {', '.join(diag_parts)} 均无数据。"
            f"可能原因: 打印机无作业, 或 25H2 安全策略限制读取。"
        )
    else:
        result["diagnostics"] = (
            f"尝试 {', '.join(diag_parts)} 均无数据, 打印机可能暂无打印作业。"
        )
    return result


def open_system_printer_queue(printer_name: str) -> bool:
    """
    打开 Windows 系统自带的打印机队列窗口（25H2 兼容）。

    使用 rundll32 printui.dll,PrintUIEntry /o /n "打印机名" 命令，
    这是微软官方支持的打开系统打印队列窗口的方式。

    参数:
        printer_name: 打印机名称

    返回: 是否成功启动
    """
    try:
        subprocess.Popen(
            ["rundll32", "printui.dll,PrintUIEntry", "/o", "/n", printer_name],
            shell=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except Exception:
        return False


def cancel_print_job(printer_name: str, job_id: int) -> bool:
    """
    取消指定的打印作业

    参数:
        printer_name: 打印机名称
        job_id: 作业ID（从 list_print_jobs 获取）

    返回: 是否成功取消
    """
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
            return True
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        return False


def cancel_all_jobs(printer_name: str = None) -> tuple:
    """
    取消打印机的所有打印作业

    参数:
        printer_name: 打印机名称，None 表示使用默认打印机

    返回: (cancelled_count, failed_count)
    """
    if printer_name is None:
        printer_name = get_default_printer()

    # list_print_jobs 返回 dict，取 "jobs" 字段
    result = list_print_jobs(printer_name)
    jobs = result.get("jobs", [])
    cancelled = 0
    failed = 0

    for job in jobs:
        if cancel_print_job(printer_name, job["job_id"]):
            cancelled += 1
        else:
            failed += 1

    return (cancelled, failed)


# ============================================================
# 打印预览 (v0.4)
# ============================================================

def preview_pdf(pdf_file: str) -> bool:
    """
    使用 SumatraPDF 或系统默认程序打开 PDF 预览

    参数:
        pdf_file: PDF 文件路径

    返回: 是否成功打开
    """
    if not os.path.exists(pdf_file):
        return False

    abs_path = os.path.abspath(pdf_file)

    # 优先使用 SumatraPDF（内置或已安装）
    sumatra_path = find_sumatra_pdf()
    if sumatra_path:
        try:
            # 使用 CREATE_NEW_PROCESS_GROUP 避免子进程被父进程的 Ctrl+C 影响
            # DETACHED_PROCESS 让 SumatraPDF 独立运行
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | 0x00000008  # DETACHED_PROCESS
            subprocess.Popen(
                [sumatra_path, abs_path],
                shell=False,
                creationflags=creationflags,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
        except Exception as e:
            # SumatraPDF 启动失败，尝试降级
            pass

    # 降级：使用系统默认程序打开
    try:
        os.startfile(abs_path)
        return True
    except Exception:
        return False


def get_pdf_info(pdf_file: str) -> dict:
    """
    获取 PDF 文件信息（用于预览展示）

    返回: {
        "path": str,           # 完整路径
        "filename": str,       # 文件名
        "page_count": int,     # 页数
        "file_size": int,      # 文件大小（字节）
        "file_size_mb": float, # 文件大小（MB）
    }
    """
    info = {
        "path": os.path.abspath(pdf_file),
        "filename": os.path.basename(pdf_file),
        "page_count": 0,
        "file_size": 0,
        "file_size_mb": 0.0,
    }

    if os.path.exists(pdf_file):
        info["file_size"] = os.path.getsize(pdf_file)
        info["file_size_mb"] = round(info["file_size"] / (1024 * 1024), 2)

    # 尝试读取页数（pypdf 可能不可用或文件损坏）
    try:
        from pypdf import PdfReader
        reader = PdfReader(pdf_file)
        info["page_count"] = len(reader.pages)
    except Exception:
        pass

    return info


def find_print_engine(preferred_engine: str = "auto") -> tuple:
    """
    根据首选引擎查找可用的打印引擎

    参数:
        preferred_engine: "auto"|"sumatra"|"acrobat"|"shell"

    返回: (engine_type, engine_path)
        engine_type: "sumatra"|"acrobat"|"shell" 或 None(指定引擎未找到)
        engine_path: 可执行文件路径或 None(shell引擎无路径)

    优先级 (auto模式): SumatraPDF > Acrobat > ShellExecute
    """
    if preferred_engine == "shell":
        return (ENGINE_SHELL, None)

    if preferred_engine in ("auto", "sumatra"):
        sumatra_path = find_sumatra_pdf()
        if sumatra_path:
            return (ENGINE_SUMATRA, sumatra_path)

    if preferred_engine in ("auto", "acrobat"):
        acrobat_path = find_acrobat_reader()
        if acrobat_path:
            return (ENGINE_ACROBAT, acrobat_path)

    if preferred_engine == "auto":
        return (ENGINE_SHELL, None)

    # 指定了引擎但未找到
    return (None, None)


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
# 页面范围与文件发现
# ============================================================

def parse_page_range(page_range_str: str, total_pages: int) -> list:
    """
    解析页面范围字符串，返回0-based页面索引列表

    支持的格式:
      - "all": 全部页面
      - "odd": 奇数页（1, 3, 5, ...）
      - "even": 偶数页（2, 4, 6, ...）
      - "1-5,8,10-12": 自定义范围（1-based输入）

    参数:
        page_range_str: 页面范围字符串
        total_pages: PDF总页数（用于边界检查）

    返回: 排序去重后的0-based页面索引列表
          空列表表示无匹配页面
    """
    if page_range_str == "all":
        return list(range(total_pages))

    if page_range_str == "odd":
        return [i for i in range(total_pages) if i % 2 == 0]  # 0-based: 0,2,4 → 1,3,5

    if page_range_str == "even":
        return [i for i in range(total_pages) if i % 2 == 1]  # 0-based: 1,3,5 → 2,4,6

    # 自定义范围解析（如 "1-5,8,10-12"）
    pages = set()
    for part in page_range_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                start, end = part.split("-", 1)
                start_idx = int(start.strip()) - 1   # 转为0-based
                end_idx = int(end.strip())            # 末尾包含，不-1
                if start_idx < 0:
                    start_idx = 0
                for idx in range(start_idx, min(end_idx, total_pages)):
                    pages.add(idx)
            except ValueError:
                continue  # 忽略格式错误的部分
        else:
            try:
                page_idx = int(part) - 1  # 转为0-based
                if 0 <= page_idx < total_pages:
                    pages.add(page_idx)
            except ValueError:
                continue  # 忽略格式错误的部分

    return sorted(pages)


def filter_pdf_pages(
    input_pdf: str,
    page_range: str,
    output_pdf: str,
) -> str:
    """
    根据页面范围过滤PDF页面，生成新的PDF文件

    参数:
        input_pdf: 输入PDF文件路径
        page_range: 页面范围字符串（传给 parse_page_range）
        output_pdf: 输出PDF文件路径

    返回: 实际使用的PDF路径
           - 如果 page_range=="all"，直接返回 input_pdf（不生成新文件）
           - 否则返回 output_pdf（已写入过滤后的页面）
    """
    if page_range == "all":
        return input_pdf

    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)
    indices = parse_page_range(page_range, total_pages)

    if not indices:
        return input_pdf  # 无匹配页面，返回原文件（调用方会报"没有页面可打印"）

    writer = PdfWriter()
    for idx in indices:
        writer.add_page(reader.pages[idx])

    with open(output_pdf, "wb") as f:
        writer.write(f)

    return output_pdf


def discover_pdf_files(input_path: str) -> list:
    """
    根据输入路径发现PDF文件

    参数:
        input_path: 文件路径或文件夹路径

    返回: PDF文件绝对路径列表（按名称排序）
           - 文件: 返回 [abs_path]
           - 文件夹: 递归扫描所有 *.pdf
           - 无结果: 返回 []
    """
    abs_path = os.path.abspath(input_path)

    if os.path.isfile(abs_path):
        return [abs_path]

    if os.path.isdir(abs_path):
        import glob
        pattern = os.path.join(abs_path, "**", "*.pdf")
        files = glob.glob(pattern, recursive=True)
        return sorted(files)

    return []


# ============================================================
# 打印
# ============================================================

def _print_with_sumatra(
    pdf_file: str, printer_name: str,
    sumatra_path: str, duplex_mode: str = "long",
) -> tuple:
    """
    使用 SumatraPDF 打印 PDF（同步）

    参数:
        pdf_file: PDF 文件绝对路径
        printer_name: 打印机名称
        sumatra_path: SumatraPDF.exe 路径
        duplex_mode: "long"|"short"|"none"

    返回: (success: bool, error_msg: str)

    特点:
        - subprocess.run 同步等待，打印完成或超时后返回
        - 双面模式通过 -print-settings 控制，不需要 DEVMODE
        - 缩放通过 -print-settings fit 控制（适应页面）
    """
    # 构建 -print-settings 参数
    settings_parts = [
        "fit",  # 缩放：适应页面
        SUMATRA_DUPLEX_MAP.get(duplex_mode, "duplexlong"),  # 双面模式
    ]
    settings_str = ",".join(settings_parts)

    cmd = [
        sumatra_path,
        "-print-to", printer_name,
        "-print-settings", settings_str,
        "-silent",  # 静默模式，不显示错误对话框
        pdf_file,
    ]

    try:
        result = subprocess.run(
            cmd, shell=False, capture_output=True, text=True, timeout=60,
        )
        if result.returncode == 0:
            return (True, "")
        error_codes = {
            2: "文件不存在或格式不支持",
            3: "文档不允许打印",
            4: "打印机不存在",
            5: "打印机驱动/设备失败",
            6: "打印被策略禁止",
        }
        error_msg = error_codes.get(
            result.returncode, f"SumatraPDF退出码: {result.returncode}"
        )
        if result.stderr:
            error_msg += f" | {result.stderr.strip()}"
        return (False, error_msg)

    except subprocess.TimeoutExpired:
        return (False, "SumatraPDF打印超时(60秒)")
    except Exception as e:
        return (False, f"SumatraPDF打印异常: {e}")


def _print_with_acrobat(
    pdf_file: str, printer_name: str, acrobat_path: str,
) -> tuple:
    """
    使用 Acrobat Reader 打印 PDF（异步）

    参数:
        pdf_file: PDF 文件绝对路径
        printer_name: 打印机名称
        acrobat_path: Acrobat 可执行文件路径

    返回: (success: bool, error_msg: str)

    特点:
        - Popen 异步启动，sleep(3) 等待发送
        - 双面通过 DEVMODE 配置（在 print_with_delay 中处理）
    """
    try:
        subprocess.Popen(
            [acrobat_path, "/t", pdf_file, printer_name],
            shell=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(3)  # 等待 Acrobat 启动并发送打印任务
        return (True, "")
    except Exception as e:
        return (False, f"Acrobat打印异常: {e}")


def _print_with_shell(pdf_file: str, printer_name: str) -> tuple:
    """
    使用 ShellExecute 打印 PDF（系统默认关联程序）

    参数:
        pdf_file: PDF 文件绝对路径
        printer_name: 打印机名称

    返回: (success: bool, error_msg: str)
    """
    try:
        win32api.ShellExecute(
            0, "print", pdf_file, f'/d:"{printer_name}"', ".", 0
        )
        time.sleep(2)  # 等待发送
        return (True, "")
    except Exception as e:
        return (False, f"ShellExecute打印异常: {e}")


def print_pdf_advanced(
    pdf_file: str,
    printer_name: str = None,
    engine_type: str = None,
    engine_path: str = None,
    duplex_mode: str = "long",
) -> bool:
    """
    使用指定的打印引擎打印 PDF

    参数:
        pdf_file: PDF 文件路径
        printer_name: 打印机名称，None 表示使用默认打印机
        engine_type: 引擎类型 "sumatra"|"acrobat"|"shell"
                     （由 print_with_delay 解析后传入）
        engine_path: 引擎可执行文件路径（shell 引擎为 None）
        duplex_mode: 双面模式（SumatraPDF 引擎使用）

    返回: 是否成功发送打印命令
    """
    if printer_name is None:
        printer_name = get_default_printer()

    abs_path = os.path.abspath(pdf_file)

    # 引擎未指定时降级到 shell
    if engine_type is None:
        engine_type = ENGINE_SHELL

    if engine_type == ENGINE_SUMATRA and engine_path:
        success, _ = _print_with_sumatra(abs_path, printer_name, engine_path, duplex_mode)
        return success
    elif engine_type == ENGINE_ACROBAT and engine_path:
        success, _ = _print_with_acrobat(abs_path, printer_name, engine_path)
        return success
    else:
        success, _ = _print_with_shell(abs_path, printer_name)
        return success


# ============================================================
# 主打印流程
# ============================================================

def print_with_delay(
    pdf_file: str,
    delay_seconds: int = 20,  # 每次打印后的等待时间（秒），仅双面模式生效
    printer_name: Optional[str] = None,
    keep_temp: bool = False,
    configure_duplex: bool = True,
    duplex_mode: str = "long",
    engine: str = "auto",
    progress_callback: ProgressCallback = None,
    cancel_event: Optional[threading.Event] = None,
    copies: int = 1,             # 打印份数（逐份打印）
    page_range: str = "all",     # 页面范围: all/odd/even/"1-5,8"
) -> dict:
    """
    主函数：分批打印PDF，每次打印后等待指定时间

    参数:
        pdf_file: 输入PDF文件路径
        delay_seconds: 每次打印后的等待时间（秒），仅双面模式生效
        printer_name: 打印机名称（None表示使用默认打印机）
        keep_temp: 是否保留临时文件
        configure_duplex: 是否自动配置双面打印
        duplex_mode: "long"=长边翻转, "short"=短边翻转, "none"=单面
        engine: 打印引擎 "auto"|"sumatra"|"acrobat"|"shell"
        progress_callback: 进度回调函数
        cancel_event: 取消事件（threading.Event）
        copies: 打印份数（默认1，逐份打印）
        page_range: 页面范围 "all"|"odd"|"even"|"1-5,8,10-12"

    返回: {"success_count": int, "total_sheets": int, "cancelled": bool}

    说明:
        - 双面模式(long/short): PDF按2页拆分，每批打印后延迟等待打印机机械复位
        - 单面模式(none): 不拆分PDF，直接打印整个文件，无延迟
        - 份数>1时，逐份打印（完成完整一份后再打下一份）
        - 页面范围非"all"时，在拆分前先过滤页面
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

    # 打印统计：记录开始时间和总页数
    start_time = time.time()
    total_pages = 0

    # 确定打印机
    if printer_name is None:
        printer_name = get_default_printer()
        _cb("info", {"message": f"使用默认打印机: {printer_name}"})
    else:
        _cb("info", {"message": f"使用打印机: {printer_name}"})

    # 显示打印机型号检测结果
    model_info = detect_printer_model(printer_name)
    if model_info["matched"]:
        _cb("info", {"message": f"检测到打印机型号: {model_info['brand'].upper()} {model_info['model']}"})
        _cb("info", {"message": f"推荐延迟: {model_info['recommended_delay']}秒 ({model_info['note']})"})
    else:
        _cb("info", {"message": f"未匹配已知型号，推荐延迟: {model_info['recommended_delay']}秒（{model_info['note']}）"})

    # 检测网络打印机
    if is_network_printer(printer_name):
        _cb("info", {"message": "检测为网络打印机，延迟已自动增加5秒"})

    # 解析打印引擎（只解析一次）
    engine_type, engine_path = find_print_engine(engine)
    if engine_type is None:
        _cb("error", {"message": f"指定的打印引擎不可用: {engine}"})
        _cb("done", {"success_count": 0, "total_sheets": 0, "cancelled": False})
        return result

    _cb("engine", {
        "engine": engine_type,
        "engine_name": ENGINE_NAMES.get(engine_type, engine_type),
        "engine_path": engine_path or "N/A",
    })

    # 配置双面打印
    # SumatraPDF 引擎通过 -print-settings 控制双面，不需要 DEVMODE 配置
    if configure_duplex and engine_type != ENGINE_SUMATRA:
        _cb("info", {"message": "配置打印机双面打印设置..."})
        success = configure_printer_duplex(printer_name, duplex_mode)
        _cb("config", {
            "printer": printer_name,
            "duplex_mode": duplex_mode,
            "success": success,
        })
    elif engine_type == ENGINE_SUMATRA:
        _cb("info", {"message": "SumatraPDF引擎，双面通过-print-settings控制，跳过DEVMODE配置"})
        _cb("config", {
            "printer": printer_name,
            "duplex_mode": duplex_mode,
            "success": True,
        })

    # 判断是否为双面打印模式（单面模式无需拆分和延迟）
    is_duplex = duplex_mode != "none"

    # 使用系统临时目录创建临时子目录
    temp_dir = tempfile.mkdtemp(prefix="pdf_duplex_")
    _cb("info", {"message": f"临时目录: {temp_dir}"})

    try:
        # 检查取消
        if _is_cancelled():
            result["cancelled"] = True
            return result

        # 页面范围过滤（在拆分之前）
        actual_pdf = pdf_file
        if page_range != "all":
            filtered_pdf = os.path.join(temp_dir, "filtered.pdf")
            actual_pdf = filter_pdf_pages(pdf_file, page_range, filtered_pdf)
            if actual_pdf == filtered_pdf:
                _cb("info", {"message": f"页面范围: {page_range}，已过滤页面"})
            # 如果返回原文件说明无匹配或page_range==all，继续使用原文件

        if is_duplex:
            # 双面模式：按2页一批拆分PDF，每批打印后延迟等待打印机机械复位
            temp_files = split_pdf_for_duplex(actual_pdf, temp_dir, progress_callback)
            reader = PdfReader(actual_pdf)
            total_pages = len(reader.pages)
        else:
            # 单面模式：无需拆分，直接打印整个PDF文件
            reader = PdfReader(actual_pdf)
            total_pages = len(reader.pages)
            temp_files = [actual_pdf]  # 直接使用(过滤后的)文件
            _cb("split_start", {"total_pages": total_pages, "sheets": 1, "is_duplex": False})
            _cb("info", {"message": f"单面打印模式，无需拆分PDF，共 {total_pages} 页直接打印"})
            _cb("split_done", {"sheets": 1, "temp_dir": "N/A (未拆分)"})

        total_sheets = len(temp_files)
        result["total_sheets"] = total_sheets

        if total_sheets == 0:
            _cb("error", {"message": "PDF没有页面可打印"})
            return result

        if is_duplex:
            _cb("info", {"message": f"开始打印，共 {total_sheets} 批次（每批2页）..."})
        else:
            _cb("info", {"message": f"开始打印，共 {total_sheets} 个文件..."})

        # 份数外层循环（逐份打印：先打完一份的所有批次，再打下一份）
        for copy_idx in range(copies):
            if _is_cancelled():
                result["cancelled"] = True
                break

            if copies > 1:
                _cb("copy_start", {
                    "copy_index": copy_idx + 1,
                    "total_copies": copies,
                })
                _cb("info", {"message": f"--- 第 {copy_idx + 1}/{copies} 份 ---"})

            # 逐批打印（内层循环）
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

                success = print_pdf_advanced(
                    temp_file, printer_name,
                    engine_type=engine_type,
                    engine_path=engine_path,
                    duplex_mode=duplex_mode,
                )

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

                # 延迟等待：仅双面模式需要，且不是最后一批
                # 单面模式无卡纸风险，直接跳过延迟
                if is_duplex and idx < total_sheets:
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

            # 份数之间的短暂停顿（让打印机准备好下一份）
            if copy_idx < copies - 1 and not _is_cancelled() and copies > 1:
                _cb("info", {"message": f"第{copy_idx + 1}份完成，准备下一份..."})
                time.sleep(2)

        # 双面模式额外等待确保最后一个文件被Reader完全读取
        # 单面模式无需此等待（SumatraPDF同步返回，单面无机械复位需求）
        if is_duplex and not _is_cancelled() and engine_type != ENGINE_SUMATRA:
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

    # 打印统计 (v0.4)
    elapsed = time.time() - start_time
    success_rate = (result["success_count"] / result["total_sheets"] * 100) if result["total_sheets"] > 0 else 0
    _cb("stats", {
        "total_pages": total_pages * copies,
        "total_batches": result["total_sheets"] * copies,
        "success_batches": result["success_count"],
        "failed_batches": result["total_sheets"] * copies - result["success_count"],
        "success_rate": success_rate,
        "elapsed_seconds": elapsed,
        "elapsed_formatted": f"{int(elapsed // 60)}分{int(elapsed % 60)}秒",
        "delay_seconds": delay_seconds if is_duplex else 0,
        "printer_name": printer_name,
        "engine_type": engine_type,
    })

    return result


# ============================================================
# CLI 默认回调
# ============================================================

def default_cli_callback(event_type: str, data: dict) -> None:
    """CLI 默认进度回调实现，将事件转为 print 输出"""

    if event_type == "info":
        print(data.get("message", ""))

    elif event_type == "engine":
        engine_name = data.get("engine_name", "")
        engine_path = data.get("engine_path", "")
        print(f"  打印引擎: {engine_name}")
        if engine_path and engine_path != "N/A":
            print(f"  引擎路径: {engine_path}")

    elif event_type == "copy_start":
        print(f"\n>>> 第 {data['copy_index']}/{data['total_copies']} 份 <<<")

    elif event_type == "split_start":
        print(f"PDF总页数: {data['total_pages']}")
        if data.get("is_duplex", True):
            print(f"需要打印 {data['sheets']} 张纸（双面）")
        else:
            print(f"单面打印，共 {data['total_pages']} 页")

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

    elif event_type == "stats":
        print(f"\n{'='*50}")
        print("打印统计")
        print(f"{'='*50}")
        print(f"  总页数:       {data['total_pages']}")
        print(f"  总批次数:     {data['total_batches']}")
        print(f"  成功批次:     {data['success_batches']}")
        print(f"  失败批次:     {data['failed_batches']}")
        print(f"  成功率:       {data['success_rate']:.1f}%")
        print(f"  总耗时:       {data['elapsed_formatted']}")
        print(f"  延迟设置:     {data['delay_seconds']}秒")
        print(f"  打印机:       {data['printer_name']}")
        engine_names = {"sumatra": "SumatraPDF", "acrobat": "Acrobat Reader", "shell": "系统默认"}
        print(f"  打印引擎:     {engine_names.get(data['engine_type'], data['engine_type'])}")
        print(f"{'='*50}")

    elif event_type == "error":
        print(f"错误: {data.get('message', '未知错误')}")
