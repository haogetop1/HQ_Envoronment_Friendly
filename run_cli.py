# -*- coding: utf-8 -*-
"""命令行入口（无 GUI 时的批处理模式，也用于自测）。

用法：
    python run_cli.py <源.xls> <模板.xlsx> <输出目录> [--prefix RR12345]

--prefix 可选。源文件名里能抓到 RR 编号时不需要它；抓不到时必须给，
否则中止（2026-08-31：不再硬编码兜底）。
"""
from __future__ import annotations

import os
import sys
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, 'src'))
sys.path.insert(0, _HERE)

from hsf.config import Settings
from hsf.pipeline import Pipeline


def main(argv: list):
    args = [a for a in argv[1:] if not a.startswith('--')]
    prefix = None
    if '--prefix' in argv:
        i = argv.index('--prefix')
        if i + 1 < len(argv):
            prefix = argv[i + 1]
    elif '-p' in argv:
        i = argv.index('-p')
        if i + 1 < len(argv):
            prefix = argv[i + 1]

    if len(args) < 3:
        print(__doc__)
        return 2
    src, tpl, out = args[0], args[1], args[2]
    t0 = time.time()

    def progress(msg, pct):
        print(f'[{pct:3d}%] {msg}', flush=True)

    p = Pipeline(src, tpl, out, Settings(), progress, prefix_override=prefix)
    res = p.run()
    print('-' * 70)
    for e in res.errors:
        print('ERROR:', e, flush=True)
    # 告警与 GUI / 日志口径一致：每条单独一行，带物料名（exe 的 --cli 模式
    # 之前就打印 WARN，复用这份实现后必须保留，不能静默吞掉）
    for r in res.records:
        for w in r.warnings:
            print(f'WARN  [{r.name}] {w}', flush=True)
    print(f'生成表 {res.n_xlsx} 份，HSF PDF {res.n_hsf} 份，MSDS PDF {res.n_msds} 份')
    print(f'耗时 {time.time() - t0:.1f}s -> {out}')
    return 1 if res.errors else 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv))
