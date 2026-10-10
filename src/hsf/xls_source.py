# -*- coding: utf-8 -*-
"""读取源 .xls：合并单元格、物料块划分、成分提取。"""
from __future__ import annotations

from dataclasses import dataclass, field

import xlrd

from .config import Settings
from .ole_extract import EmbeddedObject, is_pdf
from .util import clean_cas


@dataclass
class Component:
    row: int
    substance: str
    cas: str
    mass: float


@dataclass
class MaterialBlock:
    index: int
    row_start: int
    row_end: int            # 含
    name_raw: str
    name: str               # 去空格
    display: str            # 去重后的最终名称（可能带 -1 / -2）
    components: list = field(default_factory=list)
    hsf_objs: list = field(default_factory=list)     # EmbeddedObject
    msds_objs: list = field(default_factory=list)

    @property
    def excel_rows(self) -> str:
        return f'{self.row_start + 1}-{self.row_end + 1}'


class SourceSheet:
    def __init__(self, path: str, settings: Settings):
        self.path = path
        self.st = settings
        self.book = xlrd.open_workbook(path, formatting_info=True)
        try:
            self.sheet = self.book.sheet_by_name(settings.sheet_name)
        except xlrd.XLRDError:
            self.sheet = self.book.sheet_by_index(0)
        self.nrows = self.sheet.nrows
        self.ncols = self.sheet.ncols
        # (row, col) -> (anchor_row, anchor_col)
        self._merge_map = {}
        for r0, r1, c0, c1 in self.sheet.merged_cells:
            for r in range(r0, r1):
                for c in range(c0, c1):
                    self._merge_map[(r, c)] = (r0, c0)

    # ---------------- 单元格读取（合并感知）
    def cell(self, row: int, col: int):
        if row < 0 or row >= self.nrows or col < 0 or col >= self.ncols:
            return ''
        r, c = self._merge_map.get((row, col), (row, col))
        return self.sheet.cell_value(r, c)

    def text(self, row: int, col: int) -> str:
        v = self.cell(row, col)
        if isinstance(v, float):
            if abs(v - round(v)) < 1e-9:
                return str(int(round(v)))
            return repr(v)
        return str(v).strip()

    def number(self, row: int, col: int) -> float:
        v = self.cell(row, col)
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).strip())
        except Exception:
            return 0.0

    # ---------------- 数据区
    def detect_data_range(self):
        st = self.st
        start = st.data_start_row
        end = self.nrows - 1
        for r in range(start, self.nrows):
            v = self.text(r, 0).lower()
            if 'total' in v or '合计' in v or '总计' in v:
                end = r - 1
                break
        return start, end

    def distinct_cells(self, col: int, r0: int, r1: int):
        """返回该列在 [r0, r1] 内的「单元格」列表：(start_row, end_row)。

        合并单元格算一个，未合并的独立行各算一个。
        """
        blocks = []
        r = r0
        while r <= r1:
            anchor = self._merge_map.get((r, col))
            if anchor:
                ar, ac = anchor
                span = next((m for m in self.sheet.merged_cells
                             if m[0] == ar and m[2] == ac), None)
                if span and span[2] == col:
                    end = min(span[1] - 1, r1)
                else:
                    end = r
                blocks.append((max(ar, r0), min(end, r1)))
                r = end + 1
            else:
                blocks.append((r, r))
                r += 1
        return blocks

    # ---------------- 物料块
    def build_blocks(self, objs: list[EmbeddedObject]):
        """按 K 列合并格切块，返回 (blocks, skipped)。

        空块（不出表、不计入结果）判定，任一命中即跳过：
          A. 该 K 列单元格内没有任何成分行（F/G 列全空）
          B. 该 K 列单元格内没有内嵌 HSF 报告——没有报告就没有测试结果
             （对应需求「提取不到 PDF 就不出结果」）

        MSDS 不受影响：它按 M 列合并格独立另存，不依赖物料块。
        """
        st = self.st
        r0, r1 = self.detect_data_range()
        cells = self.distinct_cells(st.columns.hsf, r0, r1)

        blocks: list[MaterialBlock] = []
        skipped: list = []          # [(row_start, row_end, reason)]
        for (rs, re_) in cells:
            if rs > re_:
                continue
            name_raw = self.text(rs, st.columns.material_name)
            comps = []
            for r in range(rs, re_ + 1):
                sub = self.text(r, st.columns.component)
                cas = self.text(r, st.columns.cas)
                if not sub and not cas:
                    continue
                comps.append(Component(r, sub, cas, self.number(r, st.columns.mass)))
            if not comps:
                skipped.append((rs, re_, 'K 列单元格内没有成分行'))
                continue
            blocks.append(MaterialBlock(len(blocks), rs, re_,
                                        name_raw, '', '', comps))

        # PDF 归属
        for o in objs:
            if not is_pdf(o):
                continue
            for b in blocks:
                if b.row_start <= o.row <= b.row_end:
                    if o.col == st.columns.hsf:
                        b.hsf_objs.append(o)
                    elif o.col == st.columns.msds:
                        b.msds_objs.append(o)
                    break

        # 空块剔除：没有 HSF 报告的块不出表
        kept: list[MaterialBlock] = []
        for b in blocks:
            if not b.hsf_objs:
                skipped.append((b.row_start, b.row_end,
                                'K 列单元格内没有内嵌 HSF 报告'))
                continue
            kept.append(b)

        for i, b in enumerate(kept):
            b.index = i
            b.hsf_objs.sort(key=lambda o: (o.row, o.col, o.obj_id))
            b.msds_objs.sort(key=lambda o: (o.row, o.col, o.obj_id))
        skipped.sort()
        return kept, skipped


def merge_components(components: list[Component]) -> tuple:
    """同一 K 单元格内，清洗后相同的 CAS NO 合并成一行、质量相加。

    2026-08-31 需求：G46/G47 同为 66402-68-4，占比 73.15% 与 2.85%，
    应写成一行 76%。这里叠加的是**质量（H 列）**，之后再统一归一化成百分比——
    数学上等价于「先归一化再相加」，但只有一次舍入，不会累积误差。

    合并键用 clean_cas() 的结果，所以 `7440-57-5(金)` 与 `7440-57-5` 会被
    正确视为同一个 CAS。

    保密行（清洗后 CAS 为空）**不参与合并**，各自独立保留：两行都没写 CAS 时
    无法判断它们是不是同一种物质，贸然合并会丢信息。

    返回 (合并后的 Component 列表, 合并说明列表)。
    """
    order: list = []
    agg: dict = {}
    rows_by_key: dict = {}

    for c in components:
        key = clean_cas(c.cas)
        mass = max(c.mass, 0.0)
        if not key:                                  # 保密行：原样保留，不合并
            order.append(Component(c.row, c.substance, c.cas, mass))
            continue
        if key in agg:
            agg[key].mass += mass
            # 名称取第一条非空的（grill-me 决策：名称取第一条）
            if not (agg[key].substance or '').strip() and (c.substance or '').strip():
                agg[key].substance = c.substance
            rows_by_key[key].append(c.row)
        else:
            agg[key] = Component(c.row, c.substance, c.cas, mass)
            rows_by_key[key] = [c.row]
            order.append(agg[key])

    notes = []
    for key, rows in rows_by_key.items():
        if len(rows) > 1:
            cells = '、'.join(f'G{r + 1}' for r in rows)
            notes.append(f'CAS {key} 在 {cells} 共出现 {len(rows)} 次，已合并为一行'
                         f'（质量相加，物质名称取 {rows[0] + 1} 行）')
    return order, notes


def normalize_content(components: list[Component], decimals: int = 2) -> list[float]:
    """把质量归一化成百分比，并保证加总精确等于 100（最大余数法）。"""
    masses = [max(c.mass, 0.0) for c in components]
    total = sum(masses)
    n = len(masses)
    if n == 0:
        return []
    if total <= 0:                       # 没有质量数据 -> 平均分配
        base = round(100.0 / n, decimals)
        vals = [base] * n
        vals[-1] = round(100.0 - base * (n - 1), decimals)
        return vals
    scale = 10 ** decimals
    exact = [m * 100.0 / total for m in masses]
    floors = [int(e * scale) for e in exact]
    rem = int(round(100.0 * scale)) - sum(floors)
    if rem > 0:
        order = sorted(range(n), key=lambda i: (exact[i] * scale) - floors[i], reverse=True)
        for k in range(rem):
            floors[order[k % n]] += 1
    elif rem < 0:
        order = sorted(range(n), key=lambda i: (exact[i] * scale) - floors[i])
        for k in range(-rem):
            floors[order[k % n]] -= 1
    return [v / scale for v in floors]
