# -*- coding: utf-8 -*-
"""HSF / MSDS 报告的另存与合并。"""
from __future__ import annotations

import io
import os

from pypdf import PdfWriter, PdfReader

from .ole_extract import EmbeddedObject
from .util import safe_name


def merge_pdfs(datas: list[bytes]) -> bytes:
    """把多份 PDF 顺序拼接成一份；只有一份时原样返回。

    2026-08-31 修一个静默损坏数据的坑：PdfReader 必须**保活到 write() 之后**。
    add_page() 只是把页对象的引用挂到 writer 上，底层数据仍归 reader 所有；
    reader 一旦被 GC，那些页就变成空壳，写出的是一个没有内容的 PDF。
    实测某次 3 份报告（5.6/4.3/4.4 MB）合并后只剩 311 字节——
    不报错、不告警，直到打开才发现是空的。
    """
    if not datas:
        return b''
    if len(datas) == 1:
        return datas[0]
    writer = PdfWriter()
    readers: list = []                 # 保活用，千万别在 write() 之前释放
    for d in datas:
        try:
            reader = PdfReader(io.BytesIO(d))
        except Exception:
            continue
        readers.append(reader)
        for page in reader.pages:
            writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    out = buf.getvalue()
    readers.clear()                    # write() 已完成，此时释放才安全
    return out


def save_pdf_group(out_dir: str, prefix: str, kind: str, name: str,
                   objs: list[EmbeddedObject], used: set | None = None) -> tuple:
    """把一组内嵌 PDF 另存为 ``{kind}_{prefix}-{name}[-x].pdf``。

    kind 为 'HSF' 或 'MSDS'。多份时合并成一份。返回 (保存路径 or '', 告警)。

    ``used`` 是本次运行已经用过的路径集合，用来避免**批内**重名。

    2026-08-31 改：以前是 ``while os.path.exists(path)``，探测磁盘。
    那会让「重跑到同一个输出目录」时所有 PDF 都加后缀避让——
    24 份变 48 份，xlsx 却只有 24 份（xlsx 走的是覆盖写），
    三者份数对不上，而且目录里混着上一轮的陈旧文件，极易导错。
    工具应当幂等：重跑 = 重新生成，覆盖同名产物。

    批内重名仍然要避让——`_dedup()` 只在**重复计数**层面加后缀，
    挡不住「源表本来就有个 Dielectric-1」和「第二个 Dielectric 加出的
    Dielectric-1」相撞，所以这个集合不能省。
    """
    warns: list = []
    objs = [o for o in objs if o.data]
    if not objs:
        return '', warns
    if used is None:
        used = set()
    base = safe_name(f'{kind}_{prefix}-{name}', default=f'{kind}_{prefix}')
    path = os.path.join(out_dir, base + '.pdf')
    i = 1
    while path in used:                      # 只避让本次运行内已用过的名字
        path = os.path.join(out_dir, f'{base}-{i}.pdf')
        i += 1
    used.add(path)
    data = merge_pdfs([o.data for o in objs])
    if not data:
        warns.append(f'{base}.pdf 合并失败')
        return '', warns
    with open(path, 'wb') as f:
        f.write(data)
    if len(objs) > 1:
        warns.append(f'{base}.pdf 由 {len(objs)} 份报告合并而成')
    return path, warns
