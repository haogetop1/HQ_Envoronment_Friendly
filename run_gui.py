# -*- coding: utf-8 -*-
"""GUI 入口（打包成 exe 的程序入口）。

直接双击运行 -> 图形界面（支持拖拽）。
带参数运行  -> 批处理模式，便于自动化与自测：

    HSF均质表自动填写工具.exe --cli <源.xls> <模板.xlsx> <输出目录>
"""
from __future__ import annotations

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, 'src'))
sys.path.insert(0, _HERE)


def run_cli(args: list) -> int:
    """批处理模式：复用 run_cli.py 的实现。

    2026-08-31：这里原本是和 run_cli.py 并列的第二份 CLI 实现，结果加了
    --prefix 只改了那一份，exe 走的还是这份旧的，参数静默失效。
    两份 CLI 逻辑必然走偏，统一成一个入口。
    """
    from run_cli import main as cli_main
    return cli_main(['hsf-cli'] + list(args))


def main() -> int:
    argv = sys.argv[1:]
    if argv and argv[0] in ('--cli', '-c', '/cli'):
        return run_cli(argv[1:])

    # 图形模式下把 stdout/stderr 重定向掉：--windowed 打包没有控制台，
    # 任何 print 都会抛异常，不能让它干扰界面。
    if sys.executable.endswith('.exe'):
        devnull = open(os.devnull, 'w', encoding='utf-8')
        sys.stdout = devnull
        sys.stderr = devnull

    from hsf.gui import launch
    launch()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
