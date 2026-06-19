# Changelog

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

---

## 未来计划

### v0.4 (计划中)
- [ ] 添加打印队列管理（查看/取消已发送的任务）
- [ ] 支持更多打印机型号（自动检测并应用最佳延迟时间）
- [ ] 添加打印预览功能
- [ ] 支持网络打印机（延迟时间自动调整）
- [ ] 添加打印统计（总页数、耗时、成功率等）

### 已知问题
- Windows控制台中文乱码（编码问题，不影响功能）
- tkinterdnd2 为可选依赖，未安装时拖拽功能不可用
- 部分打印机驱动不支持 DEVMODE 配置，需手动设置双面模式
