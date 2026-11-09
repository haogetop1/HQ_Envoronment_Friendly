# -*- coding: utf-8 -*-
"""从 .xls（BIFF8）中提取内嵌的 OLE 对象（PDF），并定位其锚定单元格。

原理：
  1. OLE2 容器里每个嵌入对象是一个名为 ``MBD<8位十六进制>`` 的 storage，
     其中的 ``\\x01Ole10Native`` 流内含有原始文件名 + 文件正文。
  2. OBJ 的 FtPictFmla.lPosInCtlStm 给出 MBD 存储编号，与形状编号独立。
  3. OBJ 记录里不含位置；真实位置在 Escher（MSODRAWING, 0x00EC）中：
     SpContainer(0xF004) -> SP(0xF00A, 8字节: spid+grbit) + OPT(0xF00B)
     + ClientAnchor(0xF010, 18字节: flag + colL/dxL/rwT/dyT/colR/dxR/rwB/dyB)。
     顶边行 rwT 与底边行 rwB 都要留：图标压在两个合并格之间时靠底边判归属。
  4. SpContainer 的出场顺序与 OBJ 记录的出场顺序一一对应，据此配对。
"""
from __future__ import annotations

import struct
import ntpath
from dataclasses import dataclass
from itertools import takewhile

import olefile

MBD_MASK = 0x80000000


@dataclass
class EmbeddedObject:
    obj_id: int
    col: int          # 0-based 列号（锚点左上角）
    row: int          # 0-based 行号（锚点左上角 = 顶边所在行）
    file_name: str    # 内嵌时记录的原始文件名
    data: bytes       # 文件正文
    row_end: int = -1  # 0-based 行号（锚点右下角 = 底边所在行），-1 表示未知

    @property
    def bottom_row(self) -> int:
        """图标底边所在行；解析不到底边时退回顶边。"""
        return self.row_end if self.row_end >= 0 else self.row



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
        if i + 8 + cb > end:
            break
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
    for encoding in ('utf-8', 'gb18030', 'cp1252'):
        try:
            src = buf[:j].decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        src = buf[:j].decode('utf-8', 'replace')
    rest = buf[j + 1:]
    # 正文从常见文件魔数开始；找不到就整体返回
    for magic in (b'%PDF', b'PK\x03\x04', b'\xd0\xcf\x11\xe0'):
        k = rest.find(magic)
        if k >= 0:
            return ntpath.basename(src), rest[k:]
    return ntpath.basename(src), rest


def _worksheet_records(workbook: bytes, sheet_name: str):
    sheets = {}
    for rid, _, _, body in takewhile(lambda r: r[0] != 0x000A, _iter_biff_records(workbook)):
        if rid == 0x0085 and len(body) >= 8 and body[5] == 0:
            wide = body[7] & 1
            name = body[8:8 + body[6] * (2 if wide else 1)].decode('utf-16-le' if wide else 'latin1')
            sheets[name] = struct.unpack_from('<I', body)[0]
    if not sheets:
        return []
    offset = sheets.get(sheet_name, next(iter(sheets.values())))
    return list(takewhile(lambda r: r[0] != 0x000A, _iter_biff_records(workbook[offset:])))


def _storage_name(obj: bytes, obj_id: int) -> str | None:
    # MS-XLS FtPictFmla: the storage id follows the length-prefixed ObjFmla.
    fields = {rid: data for rid, _, _, data in _iter_biff_records(obj)}
    pict = fields.get(0x0009)
    if pict is None:
        return f'MBD{MBD_MASK | obj_id:08X}'
    if len(pict) < 13 or pict[8] != 2:  # PtgTbl denotes an embedded object.
        return None
    offset = 2 + struct.unpack_from('<H', pict)[0]
    if offset + 4 > len(pict):
        return None
    return f'MBD{struct.unpack_from("<I", pict, offset)[0]:08X}'


def extract_embedded_objects(xls_path: str, sheet_name='材料成分表') -> list[EmbeddedObject]:
    """返回该文件里所有内嵌对象（含锚点行列）。"""
    from .formats import workbook_format
    if workbook_format(xls_path) == 'xlsx':
        from .ooxml import extract_xlsx
        return extract_xlsx(xls_path, sheet_name)
    ole = olefile.OleFileIO(xls_path)
    try:
        entries = ole.listdir(streams=True, storages=True)

        # --- 1) OBJ 记录 -> 对象 id（按出场顺序）
        wb_name = 'Workbook' if ole.exists('Workbook') else 'Book'
        wb = ole.openstream(wb_name).read()
        records = _worksheet_records(wb, sheet_name)
        objects = []
        for rid, rlen, off, body in records:
            if rid == 0x005D and len(body) >= 10:      # OBJ
                _ot, obj_id, _gr = struct.unpack('<HHH', body[4:10])
                objects.append((obj_id, _storage_name(body, obj_id) if _ot == 8 else None))

        # --- 2) Escher -> 每个带锚点的形状
        buf = bytearray()
        for rid, rlen, off, body in records:
            if rid == 0x00EC:                          # MSODRAWING
                buf += body
        esc: list = []
        _walk_escher(bytes(buf), 0, len(buf), esc)

        anchors = []                                   # [(spid, col, row, row_end)]
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
            # ClientAnchor = flag(2B) + colL,dxL,rwT,dyT,colR,dxR,rwB,dyB（各 2B）
            # rwB 是图标底边所在行。压在两个 K 合并格之间的图标必须靠它判归属
            # （2026-10-12 需求：镍浆料（nickel paste）两份报告的顶边在第 35 行、底边在第
            # 36 行，实际属于第 36 行的 Nickel，不是 Ceramics）。
            _flag, colL, _dxL, rwT, _dyT, _colR, _dxR, rwB, _dyB = struct.unpack(
                '<9H', anchor[4])
            anchors.append((spid, colL, rwT, rwB))

        # --- 3) 按顺序配对（形状数应与 OBJ 数一致）
        streams = {e[0]: e for e in entries if len(e) == 2 and e[-1].endswith('Ole10Native')}
        if not any(storage in streams for _, storage in objects):
            return []
        if len(objects) != len(anchors):
            raise ValueError('内嵌对象与图标位置数量不一致，无法可靠确定报告归属。')

        # --- 4) 取正文
        results: list[EmbeddedObject] = []
        for (obj_id, storage), (_, col, row, row_end) in zip(objects, anchors):
            stream = streams.get(storage)
            if stream is None:
                continue
            raw = ole.openstream(stream).read()
            file_name, data = _parse_ole10native(raw)
            results.append(EmbeddedObject(obj_id, col, row, file_name, data, row_end))
        return results
    finally:
        ole.close()


def is_pdf(obj: EmbeddedObject) -> bool:
    return obj.data[:4] == b'%PDF'
