#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
下载 SumatraPDF 便携版到 vendor 目录

使用方法:
  python download_sumatra.py              # 下载默认版本
  python download_sumatra.py --version 3.5.2  # 下载指定版本
"""

import argparse
import struct
import sys
from pathlib import Path

DEFAULT_VERSION = "3.5.2"


def get_default_version():
    """获取默认版本（最新稳定版）"""
    return DEFAULT_VERSION


def download_sumatra(version, output_dir, force=False):
    """
    下载 SumatraPDF 到指定目录

    Args:
        version: 版本号，如 "3.5.2"
        output_dir: 输出目录（vendor/）

    Returns:
        bool: 成功返回 True
    """
    # 检测系统架构
    is_64bit = struct.calcsize("P") * 8 == 64
    arch_suffix = "64" if is_64bit else "32"

    # 构建下载 URL
    url = f"https://www.sumatrapdfreader.org/dl/rel/{version}/SumatraPDF-{version}-{arch_suffix}.exe"

    # 输出文件路径
    output_path = Path(output_dir) / "SumatraPDF.exe"
    
    # 如果已存在，根据 force 参数决定是否覆盖
    if output_path.exists() and not force:
        print(f"文件已存在: {output_path}")
        print("跳过下载 (使用 --force 强制覆盖)")
        return True

    print(f"正在下载 SumatraPDF {version} ({arch_suffix}位)...")
    print(f"  URL: {url}")
    print(f"  保存至: {output_path}")

    try:
        from urllib.request import urlopen, Request
        from urllib.error import URLError

        # 设置 User-Agent
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})

        with urlopen(req, timeout=30) as response:
            total_size = int(response.headers.get("Content-Length", 0))
            downloaded = 0
            block_size = 8192

            # 创建输出目录
            output_path.parent.mkdir(parents=True, exist_ok=True)

            with open(output_path, "wb") as f:
                while True:
                    buffer = response.read(block_size)
                    if not buffer:
                        break
                    downloaded += len(buffer)
                    f.write(buffer)

                    # 显示进度
                    if total_size > 0:
                        pct = int(downloaded / total_size * 100)
                        downloaded_kb = downloaded // 1024
                        total_kb = total_size // 1024
                        print(f"\r  进度: {pct}% ({downloaded_kb}KB/{total_kb}KB)", end="")

            print(f"\n下载完成: {output_path}")
            return True

    except URLError as e:
        print(f"\n下载失败 (网络错误): {e}")
        return False
    except Exception as e:
        print(f"\n下载失败: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="下载 SumatraPDF 便携版")
    parser.add_argument("--version", default=get_default_version(),
                        help=f"指定版本号 (默认: {get_default_version()})")
    parser.add_argument("--force", action="store_true",
                        help="强制覆盖已存在的文件")
    parser.add_argument("--output-dir", default="vendor",
                        help="输出目录 (默认: vendor)")

    args = parser.parse_args()

    # 获取脚本所在目录（项目根目录）
    script_dir = Path(__file__).parent
    output_dir = script_dir / args.output_dir

    success = download_sumatra(args.version, output_dir, args.force)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
