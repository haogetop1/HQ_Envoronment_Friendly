# -*- coding: utf-8 -*-
"""命令行入口（无 GUI 时的批处理模式，也用于自测）。

用法：
    python run_cli.py <源表> <模板.xlsx> <输出目录> [--prefix RR12345] [--pdf-dir 原始报告文件夹]

--prefix 可选。源文件名里能抓到 RR 编号时不需要它；抓不到时必须给，
否则中止（2026-08-31：不再硬编码兜底）。
"""
from __future__ import annotations

import os
import argparse
import sys
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, 'src'))
sys.path.insert(0, _HERE)

from hsf.config import Settings
from hsf.pipeline import Pipeline


def main(argv: list):
    parser = argparse.ArgumentParser(description='从 Excel 或 Numbers 材料成分表生成均质导入表。')
    parser.add_argument('source', help='材料成分表')
    parser.add_argument('template', nargs='?', default='', help='均质创建导入模板 .xlsx；Numbers 仅转换时可省略')
    parser.add_argument('output', help='输出目录')
    parser.add_argument('-p', '--prefix', help='手动指定芯片编号')
    parser.add_argument('--pdf-dir', help='Numbers 原始报告文件夹')
    args = parser.parse_args(argv[1:])
    t0 = time.time()

    def progress(msg, pct):
        print(f'[{pct:3d}%] {msg}', flush=True)

    p = Pipeline(args.source, args.template, args.output, Settings(), progress,
                 prefix_override=args.prefix, pdf_dir=args.pdf_dir)
    res = p.run()
    print('-' * 70)
    if res.converted_xlsx:
        print(f'已转换源表：{res.converted_xlsx}', flush=True)
    for e in res.errors:
        print('ERROR:', e, flush=True)
    # 告警与 GUI / 日志口径一致：每条单独一行，带物料名（exe 的 --cli 模式
    # 之前就打印 WARN，复用这份实现后必须保留，不能静默吞掉）
    for r in res.records:
        for w in r.warnings:
            print(f'WARN  [{r.name}] {w}', flush=True)
    print(f'生成表 {res.n_xlsx} 份，HSF PDF {res.n_hsf} 份，MSDS PDF {res.n_msds} 份')
    print(f'耗时 {time.time() - t0:.1f}s -> {args.output}')
    return 1 if res.errors else 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv))
