# -*- mode: python ; coding: utf-8 -*-
"""
PDF双面打印延迟控制脚本 - PyInstaller 打包配置
生成非单文件的可执行程序（目录模式），包含 CLI 和 GUI 两个 exe

使用方法:
  pyinstaller build.spec

输出:
  dist/pdf-duplex-printer/
    ├── pdf_duplex_printer_cli.exe   (命令行版)
    ├── pdf_duplex_printer_gui.exe   (图形界面版)
    ├── internal/                     (依赖文件)
    ├── vendor/                       (SumatraPDF便携版)
    ├── README.md
    ├── CHANGELOG.md
    └── requirements.txt
"""

from os import makedirs, pathsep
from os.path import basename, dirname, exists, join
from shutil import copyfile

from PyInstaller.utils.hooks import collect_all, get_module_file_attribute

# ============================================================
# tkinterdnd2 数据文件收集（可选依赖）
# ============================================================

def get_tkinterdnd2_datas():
    """
    收集 tkinterdnd2 的数据文件（主要是 tkdnd 文件夹）
    如果未安装 tkinterdnd2，返回空列表
    """
    try:
        import tkinterdnd2
        pkg_dir = dirname(tkinterdnd2.__file__)
        tkdnd_src = join(pkg_dir, "tkdnd")
        
        datas = []
        if exists(tkdnd_src):
            # 将 tkdnd 文件夹复制到打包后的 tkinterdnd2/tkdnd/
            datas.append((tkdnd_src, "tkinterdnd2/tkdnd"))
        
        return datas
    except ImportError:
        return []


# ============================================================
# 依赖收集
# ============================================================

binaries = []
datas = []
hiddenimports = []

# 收集 pypdf 模块的所有依赖（数据文件、二进制、隐式导入）
for module in ["pypdf"]:
    module_datas, module_binaries, module_hiddenimports = collect_all(module)
    datas += module_datas
    binaries += module_binaries
    hiddenimports += module_hiddenimports

# pywin32 的隐式导入
hiddenimports += [
    "win32print",
    "win32api",
    "win32con",
    "winreg",
    "pythoncom",
    "pywintypes",
]

# tkinter 是 Python 内置，PyInstaller 自动处理其 tcl/tk 依赖

# tkinterdnd2 可选依赖（GUI 拖拽支持）
# 如果安装了 tkinterdnd2，添加隐藏导入和数据文件
try:
    import tkinterdnd2
    hiddenimports += ["tkinterdnd2"]
    # 收集 tkinterdnd2 的数据文件（tkdnd 文件夹）
    tkinterdnd2_datas = get_tkinterdnd2_datas()
    datas += tkinterdnd2_datas
    print(f"[build.spec] tkinterdnd2 数据文件: {tkinterdnd2_datas}")
except ImportError:
    print("[build.spec] tkinterdnd2 未安装，跳过（拖拽功能将不可用）")


# ============================================================
# 排除不需要的模块（减小体积）
# ============================================================

common_excludes = [
    "IPython",
    "PIL",
    "PySide6",
    "PySide2",
    "PyQt5",
    "PyQt6",
    "matplotlib",
    "numpy",
    "pandas",
    "scipy",
    "wx",
    "notebook",
    "jupyter",
    "pytest",
    "sphinx",
]


# ============================================================
# Analysis - CLI 入口
# ============================================================

cli_analysis = Analysis(
    ["cli_app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    excludes=common_excludes,
    noarchive=False,
)


# ============================================================
# Analysis - GUI 入口
# ============================================================

gui_analysis = Analysis(
    ["gui_app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    excludes=common_excludes,
    noarchive=False,
)


# ============================================================
# PYZ (Python Zlib Archive)
# ============================================================

cli_pyz = PYZ(cli_analysis.pure)
gui_pyz = PYZ(gui_analysis.pure)


# ============================================================
# EXE - CLI (控制台程序)
# ============================================================

cli_exe = EXE(
    cli_pyz,
    cli_analysis.scripts,
    [],
    exclude_binaries=True,
    name="pdf_duplex_printer_cli",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,              # 控制台程序
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory="internal",
)


# ============================================================
# EXE - GUI (窗口程序，无控制台)
# ============================================================

gui_exe = EXE(
    gui_pyz,
    gui_analysis.scripts,
    [],
    exclude_binaries=True,
    name="pdf_duplex_printer_gui",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,             # 无控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory="internal",
)


# ============================================================
# COLLECT - 合并到统一目录（非单文件模式）
# ============================================================

coll = COLLECT(
    cli_exe,
    cli_analysis.binaries,
    cli_analysis.datas,
    gui_exe,
    gui_analysis.binaries,
    gui_analysis.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="pdf-duplex-printer",
)


# ============================================================
# 复制额外文件到 dist 目录
# ============================================================

extra_files = [
    "README.md",
    "CHANGELOG.md",
    "requirements.txt",
]

dest_root = join("dist", basename(coll.name))
for file in extra_files:
    if not exists(file):
        continue
    dest_file = join(dest_root, file)
    makedirs(dirname(dest_file), exist_ok=True)
    copyfile(file, dest_file)

# 复制 SumatraPDF 便携版到 vendor 目录
vendor_src = join(SPECPATH, "vendor", "SumatraPDF.exe")
if exists(vendor_src):
    vendor_dest = join(dest_root, "vendor", "SumatraPDF.exe")
    makedirs(dirname(vendor_dest), exist_ok=True)
    copyfile(vendor_src, vendor_dest)
    print(f"Copied SumatraPDF to {vendor_dest}")
