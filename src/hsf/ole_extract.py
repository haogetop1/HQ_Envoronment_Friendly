# -*- coding: utf-8 -*-
"""从 .xls（BIFF8）中提取内嵌的 OLE 对象（PDF），并定位其锚定单元格。

原理：
  1. OLE2 容器里每个嵌入对象是一个名为 ``MBD<8位十六进制>`` 的 storage，
     其中的 ``\\x01Ole10Native`` 流内含有原始文件名 + 文件正文。
  2. Workbook 流中的 OBJ 记录（0x005D）给出对象 id（ftCmo.idObj），
     且 ``MBD`` 名的后缀恒等于 ``0x80000000 | idObj``。
  3. OBJ 记录里不含位置；真实位置在 Escher（MSODRAWING, 0x00EC）中：
     SpContainer(0xF004) -> SP(0xF00A, 8字节: spid+grbit) + OPT(0xF00B)
     + ClientAnchor(0xF010, 18字节: flag + colL/dxL/rwT/dyT/colR/dxR/rwB/dyB)。
  4. SpContainer 的出场顺序与 OBJ 记录的出场顺序一一对应，据此配对。
"""
from __future__ import annotations

import io
import os
import struct
from dataclasses import dataclass

import olefile

MBD_MASK = 0x80000000


@dataclass
class EmbeddedObject:
    obj_id: int
    col: int          # 0-based 列号（锚点左上角）
    row: int          # 0-based 行号（锚点左上角）
    file_name: str    # 内嵌时记录的原始文件名
    data: bytes       # 文件正文


def _iter_biff_records(blob: bytes):
    pos, n = 0, len(blob)
    while pos + 4 <= n:
        rid, rlen = struct.unpack('<HH', blob[pos:pos + 4])
        if pos + 4 + rlen > n:
            break
        yield rid, rlen, pos, blob[pos + 4:pos + 4 + rlen]
        pos += 4 + rlen


def _walk_escher(buf: bytes, start: int, end: int, out: list):
    i = start
    while i + 8 <= end:
        ver_inst, fbt, cb = struct.unpack('<HHI', buf[i:i + 8])
        body = buf[i + 8:i + 8 + cb]
        out.append((i, fbt, cb, ver_inst, body))
        if fbt in (0xF002, 0xF003, 0xF004):     # 容器：递归
            _walk_escher(buf, i + 8, i + 8 + cb, out)
        i += 8 + cb


def _parse_ole10native(blob: bytes):
    """Ole10Native -> (原始文件名, 文件正文)。"""
    buf = blob[4:]
    i = buf.find(b'\x00')
    if i < 0:
        return '', b''
    buf = buf[i + 1:]
    j = buf.find(b'\x00')
    if j < 0:
        return '', b''
    src = buf[:j].decode('mbcs', 'replace')
    rest = buf[j + 1:]
    # 正文从常见文件魔数开始；找不到就整体返回
    for magic in (b'%PDF', b'PK\x03\x04', b'\xd0\xcf\x11\xe0'):
        k = rest.find(magic)
        if k >= 0:
            return os.path.basename(src), rest[k:]
    return os.path.basename(src), rest


def extract_embedded_objects(xls_path: str) -> list[EmbeddedObject]:
    """返回该文件里所有内嵌对象（含锚点行列）。"""
    ole = olefile.OleFileIO(xls_path)
    try:
        entries = ole.listdir(streams=True, storages=True)

        # --- 1) OBJ 记录 -> 对象 id（按出场顺序）
        wb_name = 'Workbook' if ole.exists('Workbook') else 'Book'
        wb = ole.openstream(wb_name).read()
        obj_ids = []
        for rid, rlen, off, body in _iter_biff_records(wb):
            if rid == 0x005D and len(body) >= 10:      # OBJ
                _ot, obj_id, _gr = struct.unpack('<HHH', body[4:10])
                obj_ids.append(obj_id)

        # --- 2) Escher -> 每个带锚点的形状
        buf = bytearray()
        for rid, rlen, off, body in _iter_biff_records(wb):
            if rid == 0x00EC:                          # MSODRAWING
                buf += body
        esc: list = []
        _walk_escher(bytes(buf), 0, len(buf), esc)

        anchors = []                                   # [(spid, col, row)]
        for i, (off, fbt, cb, vi, body) in enumerate(esc):
            if fbt != 0xF004:                          # SpContainer
                continue
            end = off + 8 + cb
            kids, j = [], i + 1
            while j < len(esc) and esc[j][0] < end:
                kids.append(esc[j])
                j += 1
            sp = next((k for k in kids if k[1] == 0xF00A), None)
            anchor = next((k for k in kids if k[1] == 0xF010), None)
            if not sp or not anchor or len(anchor[4]) != 18:
                continue
            spid = struct.unpack('<I', sp[4][:4])[0]
            _flag, colL, _dxL, rwT = struct.unpack('<4h', anchor[4][:8])
            anchors.append((spid, colL, rwT))

        # --- 3) 按顺序配对（形状数应与 OBJ 数一致）
        pairs = []
        if len(obj_ids) == len(anchors):
            pairs = list(zip(obj_ids, anchors))
        else:
            # 退化：以 spid 相对首个形状 id 的差值为对象 id 作兜底
            base = anchors[0][0] if anchors else 0
            mapping = {spid - base: (spid, c, r) for spid, c, r in anchors}
            for oid in obj_ids:
                if oid in mapping:
                    pairs.append((oid, mapping[oid]))

        # --- 4) 取正文
        results: list[EmbeddedObject] = []
        for obj_id, (spid, col, row) in pairs:
            storage = 'MBD%08X' % (MBD_MASK | obj_id)
            stream = next((e for e in entries
                           if len(e) == 2 and e[0] == storage and e[-1].endswith('Ole10Native')), None)
            if stream is None:
                continue
            try:
                raw = ole.openstream(stream).read()
            except Exception:
                continue
            file_name, data = _parse_ole10native(raw)
            results.append(EmbeddedObject(obj_id, col, row, file_name, data))
        return results
    finally:
        ole.close()


def is_pdf(obj: EmbeddedObject) -> bool:
    return obj.data[:4] == b'%PDF'
