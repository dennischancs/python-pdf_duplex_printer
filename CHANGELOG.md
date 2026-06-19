# Changelog

## v0.4 (2026-06-19)

### 新增功能

- **打印队列管理**: 
  - CLI: `-q/--queue` 查看打印队列，`--cancel JOB_ID` 取消指定作业，`--cancel all` 取消全部
  - GUI: 「打印队列」按钮，弹窗显示作业列表，支持选中取消或全部取消
- **打印机型号数据库** (`printer_models.json`):
  - 收录 Brother/HP/Canon/Epson 等常见型号及推荐延迟时间
  - 自动检测打印机型号，显示推荐延迟和备注
  - BROTHER MFC-7480D 等已知型号标记为已验证（实测数据）
  - 未匹配型号自动使用通用分类（老旧激光/新款激光/喷墨/未知）
- **网络打印机检测**: 
  - 自动检测打印机端口类型（IP/WSD/TCP/IP）
  - 网络打印机自动增加5秒额外延迟
  - 打印机信息中显示网络/本地状态
- **打印预览**:
  - CLI: `--preview` 参数，使用 SumatraPDF 或系统默认程序打开 PDF
  - GUI: 「预览」按钮，打开选中文件预览
- **打印统计**:
  - 打印完成后显示统计信息：总页数、总批次、成功率、总耗时
  - CLI 和 GUI 日志中均显示详细统计

### 修复问题

- **Windows 控制台中文乱码**: 为 `cli_app.py`、`pdf_duplex_printer.py`、`download_sumatra.py` 添加 UTF-8 控制台编码设置（`sys.stdout.reconfigure(encoding='utf-8')`）

### 优化改进

- `show_printer_capabilities()` 新增 `is_network` 和 `model_info` 字段
- `print_with_delay()` 自动显示打印机型号检测结果和网络状态
- 「查看信息」对话框显示型号识别、推荐延迟、网络状态等详细信息
- CLi 命令新增 `--queue`、`--cancel`、`--preview` 参数
- 新增 `--preview` 用法示例

### 修复 (2026-06-20)

- **打印预览内嵌化**: 原预览仅用外部 SumatraPDF 打开 PDF 文件，现实现真正的应用内嵌预览界面
  - 新增 `pdf_preview.py` 模块，使用 PyMuPDF (fitz) 将 PDF 页面渲染为图像
  - 预览窗口反映打印设置：双面模式装订边标记（长边/短边）、页面范围筛选（全部/奇数/偶数/自定义）
  - 支持翻页、跳页、缩放（0.5x~3.0x）、适合宽度、装订标记开关
  - 后台线程渲染，token 机制丢弃过期请求，避免阻塞 UI
  - PyMuPDF 不可用时自动降级为外部 SumatraPDF 打开
  - CLI `--preview` 保持原行为（无 GUI 上下文）
- **打印队列 Windows 11 25H2 兼容性**: 修复 25H2 下队列弹窗能开但作业列表始终为空的问题
  - `list_print_jobs()` 改为多级降级策略：EnumJobs level=1 → level=2 → WMI 查询
  - 返回结构改为 dict，包含 `jobs`/`method`/`diagnostics`/`error` 诊断信息
  - 新增 `open_system_printer_queue()` 用 `rundll32 printui.dll,PrintUIEntry /o /n` 打开系统队列窗口
  - GUI 队列弹窗新增「打开系统队列」按钮作为兜底，状态栏显示获取方式和诊断信息
- **新增依赖**: PyMuPDF>=1.24.0（用于 PDF 页面渲染）

### 技术细节

- 新增文件: `printer_models.json` (打印机型号数据库), `download_sumatra.py` (SumatraPDF 自动下载)
- 新增函数: `detect_printer_model()`, `get_recommended_delay()`, `is_network_printer()`, `list_print_jobs()`, `cancel_print_job()`, `cancel_all_jobs()`, `preview_pdf()`
- 新增回调事件类型: `"stats"` (打印统计)
- `print_with_delay()` 返回值保持不变，统计通过回调传递

### 文件变更
- `pdf_duplex_printer.py`: 新增7个函数 + 模型数据库加载 + 统计追踪 + 编码修复
- `cli_app.py`: 新增3个命令 + 版本号更新 + 编码修复
- `gui_app.py`: 新增2个按钮 + 队列管理弹窗 + 统计显示 + 模型信息展示
- `build.spec`: 新增编译时 SumatraPDF 自动下载
- `download_sumatra.py`: 新增 SumatraPDF 自动下载脚本
- `printer_models.json`: 新增打印机型号数据库
- `README.md`: 文档更新
- `CHANGELOG.md`: 本文件更新

---

## v0.3 (2026-06-19)

### 新增功能
- **文件/文件夹选择**: CLI和GUI都支持选择单个PDF文件或文件夹（自动扫描 `*.pdf`）
- **GUI拖拽支持**: 支持拖拽文件/文件夹到窗口（需安装 `tkinterdnd2`）
- **份数设置**: 新增份数参数（默认1），支持逐份打印
- **页面范围选择**: 支持全部/奇数/偶数/自定义范围（如 `1-5,8,10-12`）
  - 打印文件夹时只能全部页面，页面范围控件自动禁用

### 优化改进
- **延迟默认值**: 从15秒改为20秒（双面模式）
- **单面模式优化**: 
  - 不拆分PDF，直接打印整个文件
  - 无延迟等待（双面模式才需要延迟）
  - GUI中自动禁用延迟控件
- **打印引擎系统**: 
  - 三级优先级自动选择（SumatraPDF > Acrobat Reader > ShellExecute）
  - 打包内置 SumatraPDF 3.6.1 便携版（19.4MB），开箱即用
- **GUI优化**:
  - 窗口扩大到 700x750（容纳新增控件）
  - 所有14个交互控件添加 Tooltip 提示
  - 双面模式切换时联动延迟控件状态
  - 删除不必要的弹窗（打印完成/取消改为仅日志显示）

### 修复问题
- 修复单面模式下的延迟逻辑错误（原代码 `if idx < total_sheets or True` 永远为真）
- 修复 Tooltip 绑定到 BooleanVar 而非控件的 bug
- 修复引擎解析效率问题（每批次调用改为开头解析一次）

### 技术细节
- 新增函数：`parse_page_range()`, `filter_pdf_pages()`, `discover_pdf_files()`
- `print_with_delay()` 新增参数：`copies`(份数), `page_range`(页面范围)
- 份数通过外层循环实现（自然逐份打印）
- 页面范围在 PDF 拆分前过滤（保持 `split_pdf_for_duplex` 单一职责）

### 文件变更
- `pdf_duplex_printer.py`: 核心模块（新增3个函数 + 签名扩展 + 内部逻辑重写）
- `cli_app.py`: CLI入口（新增2个参数 + 文件夹支持 + 多文件循环）
- `gui_app.py`: GUI入口（窗口扩大 + 新增6个控件 + 拖拽支持 + 联动逻辑）
- `build.spec`: 打包配置（更新输出结构说明）
- `README.md`: 文档更新（打印引擎系统、页面范围、份数说明）
- `CHANGELOG.md`: 本文件（新增）

---

## v0.2 (2026-06-19)

### 新增功能
- **多打印引擎系统**: 引入三级优先级自动选择机制
  - SumatraPDF（推荐，同步打印，有退出码）
  - Acrobat Reader（异步打印，需等待3秒）
  - 系统默认（ShellExecute，降级方案）
- **SumatraPDF便携版**: 打包内置 SumatraPDF 3.6.1（19.4MB），无需预装任何PDF阅读器

### 优化改进
- **GUI Tooltip**: 自定义 Tooltip 类（纯 tkinter），为全部14个交互控件添加悬停提示
- **引擎解析优化**: 在 `print_with_delay()` 开头解析引擎一次，避免每批次重复查找
- **DEVMODE判断修正**: SumatraPDF通过命令行参数控制双面，跳过 DEVMODE 配置

### 技术细节
- 新增引擎常量：`ENGINE_SUMATRA`, `ENGINE_ACROBAT`, `ENGINE_SHELL`
- 新增函数：`find_sumatra_pdf()`, `find_print_engine()`, `_print_with_sumatra()`, `_print_with_acrobat()`, `_print_with_shell()`
- `_get_app_dir()`: 兼容开发环境和PyInstaller打包环境定位vendor目录

---

## v0.1 (2026-06-18)

### 初始版本
- **核心功能**: PDF双面打印延迟控制，解决Brother MFC-7480D连续双面打印卡纸问题
- **单面/双面模式**: 支持长边翻转、短边翻转、单面打印
- **延迟控制**: 每批打印后等待指定时间（默认15秒），让打印机机械结构复位
- **CLI模式**: 命令行界面，支持参数配置
- **GUI模式**: 图形界面（tkinter），支持可视化操作
- **打印引擎**: 使用 Acrobat Reader 作为默认打印引擎
- **PyInstaller打包**: 支持打包为独立可执行文件

### 技术实现
- 使用 `pypdf` 库按页拆分PDF（每2页一组，对应一张纸的正反面）
- 使用 `win32print` 修改打印机驱动的 `pDevMode.Duplex` 属性配置双面打印
- 使用 `subprocess.Popen` 异步调用 Acrobat Reader 打印
- 批次间延迟等待（逐秒倒计时，支持取消）
- 临时文件自动清理
