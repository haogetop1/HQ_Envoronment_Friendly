# -*- coding: utf-8 -*-
"""读取源 .xls：合并单元格、物料块划分、成分提取。"""
from __future__ import annotations

from dataclasses import dataclass, field

import xlrd

from .config import Settings
from .ole_extract import EmbeddedObject, is_pdf
from .util import clean_cas, material_display, remove_spaces


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
    # 2026-10-12：由「同一个 M 格内同名同 CAS」的多个 K 块合并而来时，
    # 这里记每个成员的原始行区间；单块时为空列表。
    merged_rows: list = field(default_factory=list)
    # 报告图标跨两个 K 格时，归属变更的说明（接收方与失去方各记一条）
    attr_notes: list = field(default_factory=list)

    @property
    def is_merged(self) -> bool:
        return len(self.merged_rows) > 1

    def _label(self, a: int, b: int) -> str:
        return f'{a + 1}-{b + 1}' if b > a else str(a + 1)

    @property
    def excel_rows(self) -> str:
        if self.merged_rows:
            return '、'.join(self._label(a, b) for a, b in self.merged_rows)
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

    # ---------------- 名称
    def material_name(self, block: MaterialBlock) -> str:
        """材料名称主体（不含 RR 前缀），取自设置指定的那一列。

        默认 ``name_source='model'`` = **D 列（材料名称/型号）**，也是
        「同名同 CAS 合并」用来判同名的那一列。D 列里的括号内容不取
        （Die1(EMMJR68) -> Die1），与输出表 B2 及文件名保持一致。
        """
        st = self.st
        src = (st.name_source or 'model').lower()
        model = material_display(self.text(block.row_start, st.columns.material_model))
        vendor = material_display(self.text(block.row_start, st.columns.material_name))
        if src == 'vendor':
            return vendor or model
        if src == 'both':
            return material_display(f'{vendor}-{model}')
        return model or vendor

    # ---------------- 物料块
    def build_blocks(self, objs: list[EmbeddedObject], m_anchor=None):
        """按 K 列合并格切块，返回 (blocks, skipped)。

        空块（不出表、不计入结果）判定，任一命中即跳过：
          A. 该 K 列单元格内没有任何成分行（F/G 列全空）
          B. 该 K 列单元格内没有内嵌 HSF 报告——没有报告就没有测试结果
             （对应需求「提取不到 PDF 就不出结果」）

        ``m_anchor``：callable(row) -> 该行所属 M 列合并格的起始行。
        给了它就启用「同名同 CAS 合并」（见 ``_merge_same_material``）：
        合并必须发生在空块剔除**之前**，否则像 RR12347 的 D36 那样
        「自己那格没有报告、但和隔壁同名同 CAS」的行会在合并前就被丢掉，
        重量凭空少算。

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
            b = MaterialBlock(len(blocks), rs, re_, name_raw, '', '', comps)
            b.name = self.material_name(b)
            blocks.append(b)

        self._assign_pdfs(blocks, objs)

        if m_anchor is not None:
            blocks = self._merge_same_material(blocks, m_anchor)

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

    # ---------------- 内嵌 PDF 归属
    def _assign_pdfs(self, blocks: list[MaterialBlock],
                     objs: list[EmbeddedObject]):
        """把内嵌 PDF 挂到所属 K 块上。

        归属行默认取图标**底边**所在行（Settings.pdf_row_anchor='bottom'）。
        2026-10-12：图标是浮的，压在两格之间时只看顶边会挂错——
        RR12347 的两份镍浆料两份报告（Sample Name 就是 Nickel paste）
        顶边在第 35 行（Ceramics 的 K34:K35）、底边在第 36 行（Nickel），
        按顶边会把镍的报告算给陶瓷。跨格时必须留痕，别静默搬走。
        """
        st = self.st
        bottom = (getattr(st, 'pdf_row_anchor', 'bottom') or 'bottom').lower() != 'top'

        def owner(row: int):
            for b in blocks:
                if b.row_start <= row <= b.row_end:
                    return b
            return None

        def label(b: MaterialBlock) -> str:
            nm = self.text(b.row_start, st.columns.material_model).strip()
            return f'{b.row_start + 1} 行（{nm}）' if nm else f'{b.row_start + 1} 行'

        for o in objs:
            if not is_pdf(o):
                continue
            anchor_row = o.bottom_row if bottom else o.row
            target = owner(anchor_row)
            if target is None:
                continue
            if o.col == st.columns.hsf:
                target.hsf_objs.append(o)
            elif o.col == st.columns.msds:
                target.msds_objs.append(o)
            else:
                continue

            top = owner(o.row)
            if top is None or top is target:
                continue
            target.attr_notes.append(
                f'报告 {o.file_name} 的图标压在 {top.row_start + 1}-'
                f'{top.row_end + 1} 行与 {target.row_start + 1} 行之间，'
                f'已按图标底边归入 {label(target)}')
            top.attr_notes.append(
                f'报告 {o.file_name} 的图标压在本格（{top.row_start + 1}-'
                f'{top.row_end + 1} 行）与下一格之间，已按图标底边归入 '
                f'{label(target)}，本格不计入')

    # ---------------- 同名同 CAS 合并
    def _merge_same_material(self, blocks: list[MaterialBlock],
                             m_anchor) -> list[MaterialBlock]:
        """同一个 M 列合并格内，D 列名称相同且 G 列 CAS 序列相同的块合并成一块。

        2026-10-12 需求：RR12347 的 D36/D38 两行都叫 Nickel、G 列同为
        7440-02-0，本应是一份均质表。旧逻辑当成两个材料出表，而 D36 因为
        自己那格没有报告被整块丢弃——0.02142 的重量直接消失。

        判同名的列就是输出表 B2 用的那一列（Settings.name_source，默认 D 列
        「材料名称/型号」）。**不能拿 C 列（材料厂家）当键**——C 列常常是横跨
        整段的合并格（整段都是同一个厂家名），拿它当键等于不判同名。

        合并口径：
          - 成分：两块成分行拼接，同 CAS 的重量**相加**，再整体归一化到 100%
          - 报告：两块的 HSF 全并进来，合并了几份就写几个报告编号
          - MSDS：仍跟随所属 M 格，只出一份（与 HSF 一一对应）
          - 名称：取第一块的（后续 _dedup 不会再给它加 -1）

        要求 CAS 序列**完全一致**才算同一种材料。名字相同但成分不同的
        保持各自出表（名称照旧加 -1 去重），不能硬合。
        """
        grouped: dict = {}
        order: list = []
        for b in blocks:
            key = (m_anchor(b.row_start),
                   remove_spaces(b.name or b.name_raw or ''),
                   tuple(clean_cas(c.cas) for c in b.components))
            if key not in grouped:
                grouped[key] = []
                order.append(key)
            grouped[key].append(b)

        out: list[MaterialBlock] = []
        for key in order:
            members = grouped[key]
            out.append(members[0] if len(members) == 1 else self._merge_blocks(members))
        return out

    def _merge_blocks(self, members: list[MaterialBlock]) -> MaterialBlock:
        first = members[0]
        comps, hsf, msds, notes = [], [], [], []
        for b in members:
            comps.extend(b.components)
            hsf.extend(b.hsf_objs)
            msds.extend(b.msds_objs)
            notes.extend(b.attr_notes)
        merged = MaterialBlock(
            index=first.index,
            row_start=min(b.row_start for b in members),
            row_end=max(b.row_end for b in members),
            name_raw=first.name_raw, name=first.name, display=first.display,
            components=comps, hsf_objs=hsf, msds_objs=msds)
        merged.merged_rows = [(b.row_start, b.row_end) for b in members]
        merged.attr_notes = notes
        return merged


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
