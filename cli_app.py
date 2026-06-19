#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF双面打印延迟控制脚本 - CLI 入口
解决Brother MFC-7480D等打印机连续双面打印卡纸问题

延迟说明:
  - 双面模式(long/short): PDF按2页拆分，每批打印后延迟等待打印机机械复位
  - 单面模式(none): 不拆分PDF，直接打印整个文件，无延迟

用法示例:
  python cli_app.py document.pdf                    使用默认打印机，20秒延迟（双面）
  python cli_app.py document.pdf -d 20              20秒延迟
  python cli_app.py document.pdf -p Brother         模糊匹配打印机
  python cli_app.py document.pdf -p 1               按编号选择打印机
  python cli_app.py document.pdf --duplex short     短边翻转双面
  python cli_app.py document.pdf --duplex none      单面打印（无需延迟）
  python cli_app.py document.pdf -n 2               打印2份（逐份）
  python cli_app.py document.pdf -r "1-5,8"         打印第1-5页和第8页
  python cli_app.py document.pdf -r odd             只打印奇数页
  python cli_app.py ./pdf_folder/                   打印文件夹中所有PDF
  python cli_app.py -l                              列出打印机
  python cli_app.py -i -p Brother                   查看指定打印机信息
"""

__version__ = "v0.4"


import argparse
import os
import sys

# Fix Windows console encoding for Chinese characters
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from pdf_duplex_printer import (
    get_available_printers,
    get_default_printer,
    get_duplex_name,
    get_paper_size_name,
    resolve_printer,
    interactive_select_printer,
    show_printer_capabilities,
    print_with_delay,
    discover_pdf_files,
    default_cli_callback,
    list_print_jobs,
    cancel_print_job,
    cancel_all_jobs,
    preview_pdf,
    get_pdf_info,
)


def cmd_list_printers():
    """列出所有可用打印机"""
    printers = get_available_printers()
    default_printer = get_default_printer()

    print("\n可用的打印机:")
    print("-" * 50)
    for i, printer in enumerate(printers, 1):
        marker = " (默认)" if printer == default_printer else ""
        print(f"  {i}. {printer}{marker}")
    print("-" * 50)
    print(f"共 {len(printers)} 台打印机\n")


def cmd_show_printer_info(printer_input: str = None):
    """显示打印机当前配置信息"""
    if printer_input:
        matched, matches = resolve_printer(printer_input)
        if not matched:
            if matches:
                print(f"找到多个匹配的打印机:")
                for i, p in enumerate(matches, 1):
                    print(f"  {i}. {p}")
                print("请使用更精确的名称或编号。")
            else:
                print(f"未找到匹配的打印机: {printer_input}")
                print("使用 -l 查看可用打印机列表")
            return
        printer_name = matched
    else:
        printer_name = get_default_printer()

    info = show_printer_capabilities(printer_name)
    print(f"\n打印机 '{info['name']}' 当前设置:")
    print("-" * 50)
    print(f"  双面打印: {info['duplex']}")
    print(f"  纸张大小: {info['paper_size']}")
    print(f"  方向:     {info['orientation']}")
    if info.get("is_network"):
        print(f"  类型:     网络打印机（延迟会自动增加）")
    else:
        print(f"  类型:     本地打印机")

    # 型号检测信息
    model_info = info.get("model_info", {})
    if model_info.get("matched"):
        print(f"  型号:     {model_info['brand'].upper()} {model_info['model']}")
        print(f"  推荐延迟: {model_info['recommended_delay']}秒 ({model_info['note']})")
        if model_info.get("verified"):
            print(f"  数据来源: 已验证（实测）")
        else:
            print(f"  数据来源: 通用建议")
    else:
        print(f"  推荐延迟: {model_info.get('recommended_delay', '未知')}秒 ({model_info.get('note', '未知型号')})")

    if "error" in info:
        print(f"  错误:     {info['error']}")
    print("-" * 50 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="PDF双面打印延迟控制脚本 - 解决连续双面打印卡纸问题",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s document.pdf                    使用默认打印机，20秒延迟（双面）
  %(prog)s document.pdf -d 20              20秒延迟
  %(prog)s document.pdf -p Brother         模糊匹配打印机
  %(prog)s document.pdf -p 1               按编号选择打印机
  %(prog)s document.pdf --duplex short     短边翻转双面
  %(prog)s document.pdf --duplex none      单面打印（无需延迟，直接打印整个文件）
  %(prog)s document.pdf -n 2               打印2份（逐份打印）
  %(prog)s document.pdf -r "1-5,8"         自定义页面范围
  %(prog)s document.pdf -r odd             仅奇数页
  %(prog)s ./pdf_folder/                   打印文件夹中所有PDF（全部页面）
  %(prog)s -l                              列出打印机
  %(prog)s -i -p Brother                   查看指定打印机信息
        """,
    )

    parser.add_argument(
        "pdf_file", nargs="?", default=None,
        help="要打印的PDF文件路径或文件夹路径"
    )
    parser.add_argument(
        "-d", "--delay", type=int, default=20, metavar="N",
        help="每批打印后的延迟秒数（默认: 20），仅双面模式生效，单面模式自动跳过"
    )
    parser.add_argument(
        "-p", "--printer", metavar="NAME",
        help="打印机名称（支持模糊匹配如 'Brother'；或编号如 '1'）"
    )
    parser.add_argument(
        "-l", "--list", action="store_true", dest="list_printers",
        help="列出所有可用打印机"
    )
    parser.add_argument(
        "-i", "--info", action="store_true",
        help="显示打印机当前配置信息"
    )
    parser.add_argument(
        "--duplex", choices=["long", "short", "none"], default="long",
        help="双面打印模式: long=长边翻转(默认), short=短边翻转, none=单面"
    )
    parser.add_argument(
        "-n", "--copies", type=int, default=1, metavar="N",
        help="打印份数（默认: 1），逐份打印"
    )
    parser.add_argument(
        "-r", "--pages", type=str, default="all",
        help='页面范围: all=全部(默认), odd=奇数页, even=偶数页, 或自定义如 "1-5,8,10-12"。文件夹模式强制为全部'
    )
    parser.add_argument(
        "--keep-temp", action="store_true",
        help="保留临时文件（用于调试）"
    )
    parser.add_argument(
        "--no-config", action="store_true",
        help="跳过自动双面打印配置"
    )
    parser.add_argument(
        "-e", "--engine",
        choices=["auto", "sumatra", "acrobat", "shell"], default="auto",
        help="打印引擎: auto=自动选择(默认), sumatra=SumatraPDF, acrobat=Acrobat Reader, shell=系统默认"
    )
    parser.add_argument(
        "-q", "--queue", action="store_true",
        help="查看指定打印机的打印队列（需配合 -p 指定打印机，否则查看默认打印机）"
    )
    parser.add_argument(
        "--cancel", type=str, metavar="JOB_ID",
        help="取消指定打印作业（需配合 -p 指定打印机）。使用 'all' 取消所有作业"
    )
    parser.add_argument(
        "--preview", action="store_true",
        help="预览 PDF 文件（使用 SumatraPDF 或系统默认程序打开）"
    )

    args = parser.parse_args()

    # --- 列出打印机 ---
    if args.list_printers:
        cmd_list_printers()
        return

    # --- 查看打印机信息 ---
    if args.info:
        cmd_show_printer_info(args.printer)
        return

    # --- 查看打印队列 ---
    if args.queue:
        # 解析打印机
        if args.printer:
            matched, _ = resolve_printer(args.printer)
            if not matched:
                print(f"未找到匹配的打印机: {args.printer}")
                return
            printer_name = matched
        else:
            printer_name = get_default_printer()
            print(f"默认打印机: {printer_name}")

        result = list_print_jobs(printer_name)
        jobs = result.get("jobs", [])
        method = result.get("method", "none")
        diagnostics = result.get("diagnostics", "")
        error = result.get("error", "")
        print(f"\n打印队列 '{printer_name}':")
        print("-" * 60)
        if not jobs:
            print("  (队列为空)")
            if diagnostics:
                print(f"  诊断: {diagnostics}")
            if error:
                print(f"  错误: {error}")
        else:
            for job in jobs:
                size_str = f"{job['size'] / 1024:.1f}KB" if job['size'] > 0 else "-"
                print(f"  作业ID: {job['job_id']}")
                print(f"    文档:   {job['document']}")
                print(f"    状态:   {job['status']}")
                print(f"    页数:   {job['pages']}")
                print(f"    大小:   {size_str}")
                print(f"    提交者: {job['owner']}")
                if job['submitted']:
                    print(f"    时间:   {job['submitted']}")
                print()
        print("-" * 60)
        print(f"共 {len(jobs)} 个待处理作业  [获取方式: {method}]")
        if diagnostics and jobs:
            print(f"诊断: {diagnostics}")
        print()
        return

    # --- 取消打印作业 ---
    if args.cancel is not None:
        # 解析打印机
        if args.printer:
            matched, _ = resolve_printer(args.printer)
            if not matched:
                print(f"未找到匹配的打印机: {args.printer}")
                return
            printer_name = matched
        else:
            printer_name = get_default_printer()

        if args.cancel.lower() == 'all':
            print(f"\n正在取消 '{printer_name}' 的所有打印作业...")
            cancelled, failed = cancel_all_jobs(printer_name)
            print(f"已取消 {cancelled} 个，失败 {failed} 个")
        else:
            try:
                job_id = int(args.cancel)
                print(f"\n正在取消 '{printer_name}' 的作业 #{job_id}...")
                if cancel_print_job(printer_name, job_id):
                    print("取消成功")
                else:
                    print("取消失败（作业不存在或权限不足）")
            except ValueError:
                print(f"无效的作业ID: {args.cancel}（应为数字或 'all'）")
        return

    # --- 预览 PDF ---
    if args.preview:
        if not args.pdf_file:
            print("错误: 请指定要预览的 PDF 文件")
            return
        if not os.path.exists(args.pdf_file):
            print(f"错误: 找不到文件 '{args.pdf_file}'")
            return

        # 显示文件信息
        info = get_pdf_info(args.pdf_file)
        print(f"\nPDF 文件信息:")
        print("-" * 40)
        print(f"  文件名:    {info['filename']}")
        print(f"  页数:      {info['page_count']} 页")
        print(f"  文件大小:  {info['file_size']:,} 字节 ({info['file_size_mb']} MB)")
        print(f"  路径:      {info['path']}")
        print("-" * 40)

        print(f"\n正在打开预览: {info['filename']}")
        if not preview_pdf(args.pdf_file):
            print("预览失败: 无法打开 PDF 文件")
        else:
            print("预览已打开 (SumatraPDF)")
        return

    # --- 打印流程 ---
    # 检查输入路径
    if not args.pdf_file:
        parser.print_help()
        print("\n错误: 请指定要打印的PDF文件或文件夹")
        sys.exit(1)

    if not os.path.exists(args.pdf_file):
        print(f"错误: 找不到文件或文件夹 '{args.pdf_file}'")
        sys.exit(1)

    # 发现PDF文件
    is_folder = os.path.isdir(args.pdf_file)
    pdf_files = discover_pdf_files(args.pdf_file)

    if not pdf_files:
        if is_folder:
            print(f"错误: 文件夹 '{args.pdf_file}' 中没有找到PDF文件")
        else:
            print(f"错误: 不是有效的PDF文件: '{args.pdf_file}'")
        sys.exit(1)

    # 文件夹模式强制全部页面
    effective_pages = args.pages
    if is_folder:
        effective_pages = "all"

    # 解析打印机
    printer_name = None
    if args.printer:
        matched, matches = resolve_printer(args.printer)
        if matched:
            printer_name = matched
        elif matches:
            # 多匹配，交互选择
            print(f"找到多个匹配的打印机:")
            printer_name = interactive_select_printer()
            if not printer_name:
                print("未选择打印机，退出。")
                sys.exit(0)
        else:
            print(f"错误: 未找到匹配的打印机 '{args.printer}'")
            print("使用 -l 查看可用打印机列表")
            sys.exit(1)
    else:
        # 未指定打印机，询问是否使用默认
        default_p = get_default_printer()
        print(f"\n当前默认打印机: {default_p}")
        try:
            choice = input("使用默认打印机? (回车=是, n=选择其他): ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            choice = ""
        if choice.lower() == 'n':
            printer_name = interactive_select_printer()
            if not printer_name:
                print("未选择打印机，退出。")
                sys.exit(0)
        else:
            printer_name = default_p

    # 打印配置摘要
    page_range_display = {
        "all": "全部", "odd": "奇数页", "even": "偶数页",
    }.get(effective_pages, effective_pages)
    if is_folder:
        page_range_display += " (文件夹模式)"

    mode_names = {"long": "长边翻转", "short": "短边翻转", "none": "单面"}
    engine_names = {"auto": "自动选择", "sumatra": "SumatraPDF", "acrobat": "Acrobat Reader", "shell": "系统默认"}

    print("\n" + "=" * 60)
    print("PDF双面打印延迟控制程序")
    print("=" * 60)
    if len(pdf_files) == 1:
        print(f"  PDF文件:   {pdf_files[0]}")
    else:
        print(f"  PDF文件:   {args.pdf_file} ({len(pdf_files)} 个PDF文件)")
    print(f"  打印机:     {printer_name}")
    if args.duplex == "none":
        print(f"  延迟间隔:   无需 (单面打印)")
    else:
        print(f"  延迟间隔:   {args.delay}秒")
    print(f"  双面模式:   {mode_names[args.duplex]}")
    print(f"  打印份数:   {args.copies} 份（逐份打印）")
    print(f"  页面范围:   {page_range_display}")
    print(f"  打印引擎:   {engine_names[args.engine]}")
    print("=" * 60 + "\n")

    # 执行打印
    try:
        total_result = {"success_count": 0, "total_sheets": 0, "cancelled": False}

        for fidx, pdf_path in enumerate(pdf_files, 1):
            if len(pdf_files) > 1:
                print(f"[{fidx}/{len(pdf_files)}] 正在打印: {os.path.basename(pdf_path)}")
                print("-" * 40)

            result = print_with_delay(
                pdf_file=pdf_path,
                delay_seconds=args.delay,
                printer_name=printer_name,
                keep_temp=args.keep_temp,
                configure_duplex=not args.no_config,
                duplex_mode=args.duplex,
                engine=args.engine,
                copies=args.copies,
                page_range=effective_pages,
                progress_callback=default_cli_callback,
            )

            total_result["success_count"] += result["success_count"]
            total_result["total_sheets"] += result["total_sheets"]
            if result["cancelled"]:
                total_result["cancelled"] = True
                break

            if len(pdf_files) > 1 and fidx < len(pdf_files):
                print("-" * 40)

        if len(pdf_files) > 1:
            print(f"\n全部完成! 共 {len(pdf_files)} 个文件, "
                  f"成功 {total_result['success_count']}/{total_result['total_sheets']} 批次")

        if total_result["cancelled"]:
            sys.exit(130)  # 130 = 被中断
        elif total_result["success_count"] == 0:
            sys.exit(1)

    except KeyboardInterrupt:
        print("\n\n用户中断操作")
        sys.exit(130)
    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
