# -*- coding: utf-8 -*-
"""按模板生成填写完成的「均质创建导入模板」.xlsx。

模板固定结构（1-based 行号，插入前的原始位置）：
    1        说明（合并行）
    2/3      基本信息（标签在 A/C/E/G/I 列，值在 B/D/F/H/J 列）
    4        「均质成分」标题（合并行）
    5        成分表头
    6..11    6 个成分行位
    12       守护行「该行不允许删除」
    13       「HSF测试结果」标题（合并行）
    14       测试项目行（B..AC 共 29 项）
    15       测试结果行
    16       守护行
    17       报告表头（A序号/B机构/C编号/D日期）
    18..21   4 个报告行位
    22       守护行

成分行超过 6 条时在第 12 行前插行，报告行超过 4 条时在第 22 行前插行，
守护行始终被推到数据区末尾，保证「该行不允许删除」不被覆盖。
"""
from __future__ import annotations

import copy
import warnings
from dataclasses import dataclass, field
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

from .config import RESULT_ITEMS, SECRET_CAS, SECRET_CAS_VALUE, Settings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

RED_FILL = PatternFill('solid', fgColor='FFC7CE')

# 模板基本信息区：列 -> (行, 值列偏移)
_INFO_ROW_NAME = 2
_INFO_ROW_CAT = 3


@dataclass
class ComponentRow:
    substance: str = ''
    cas: str = ''
    content: float = 0.0
    secret: bool = False


@dataclass
class ReportRow:
    agency: str = ''
    report_no: str = ''
    report_date: str = ''


@dataclass
class MaterialRecord:
    """一个 K 列合并格 = 一份输出表。"""
    index: int = 0
    name: str = ''                     # 去空格后的材料名称
    display: str = ''                  # 去重后的最终名称（可能带 -1 / -2）
    excel_rows: str = ''
    components: list = field(default_factory=list)     # [ComponentRow]
    results: dict = field(default_factory=dict)        # item -> 'ND' / '217'
    reports: list = field(default_factory=list)        # [ReportRow]
    out_xlsx: str = ''                                 # 输出文件路径
    msds_out: str = ''                                 # 2026-08-31：本块对应的 MSDS PDF
    warnings: list = field(default_factory=list)       # [str]
    inferred: list = field(default_factory=list)       # 由领域规则补 ND 的项
    conflicts: list = field(default_factory=list)      # 多份报告取值冲突的项

    @property
    def file_stem(self) -> str:
        return self.display or self.name


def _is_secret(cas: str, secret_when_empty: bool = True) -> bool:
    low = (cas or '').strip().lower()
    if not low:
        return secret_when_empty
    return any(k in low for k in SECRET_CAS)


# ---------------------------------------------------------------- 行样式复制

def _copy_row_style(ws, src_row: int, dst_row: int):
    for col in range(1, ws.max_column + 1):
        s = ws.cell(row=src_row, column=col)
        d = ws.cell(row=dst_row, column=col)
        d._style = copy.copy(s._style)


def _shift_merges(ws, from_row: int, k: int):
    """把 from_row 及其之后的合并区整体下移 k 行。

    openpyxl 的 `insert_rows()` 只搬单元格，不搬合并区。模板里
    `A13:AC13`（「HSF测试结果」标题行）插行后仍死死指向第 13 行，
    而第 13 行此时已经是第 8 条成分行——该行 B..AC 列的值会被合并区吞掉，
    Excel 打开只看到序号、看不到物质名称和含量，E 列也就凑不满 100%。
    """
    if k <= 0:
        return
    new_ranges = MultiCellRange()
    for r in ws.merged_cells.ranges:
        if r.min_row >= from_row:
            new_ranges.add(CellRange(min_col=r.min_col, min_row=r.min_row + k,
                                     max_col=r.max_col, max_row=r.max_row + k))
        else:
            new_ranges.add(r)
    ws.merged_cells = new_ranges


# ---------------------------------------------------------------- 主函数

def build_workbook(template_path: str, out_path: str, rec: MaterialRecord,
                   st: Settings) -> list:
    """生成一份填好的模板，返回告警列表。"""
    warns: list = []
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]              # 模板第一张表即「模板_批量导入均质信息」

    n_comp = len(rec.components)
    n_rep = max(len(rec.reports), 1)

    k_comp = max(0, n_comp - st.comp_capacity)
    k_rep = max(0, n_rep - st.report_capacity)

    # --- 先给成分行腾位置（在第 12 行守护行之前插）
    if k_comp:
        ws.insert_rows(st.tpl_comp_marker, k_comp)
        _shift_merges(ws, st.tpl_comp_marker, k_comp)
        for r in range(st.tpl_comp_marker, st.tpl_comp_marker + k_comp):
            _copy_row_style(ws, st.tpl_comp_marker - 1, r)

    guard_rep = st.tpl_report_marker + k_comp
    if k_rep:
        ws.insert_rows(guard_rep, k_rep)
        _shift_merges(ws, guard_rep, k_rep)
        for r in range(guard_rep, guard_rep + k_rep):
            _copy_row_style(ws, guard_rep - 1, r)

    # --- 基本信息
    ws.cell(row=_INFO_ROW_NAME, column=2, value=rec.name)
    ws.cell(row=_INFO_ROW_CAT, column=2, value=st.category)
    ws.cell(row=_INFO_ROW_NAME, column=4, value=st.supplier)
    ws.cell(row=_INFO_ROW_NAME, column=6, value=st.model)
    ws.cell(row=_INFO_ROW_NAME, column=8, value=st.eco_level)
    ws.cell(row=_INFO_ROW_NAME, column=10, value=None)     # 回收料添加比例 留空
    ws.cell(row=_INFO_ROW_CAT, column=4, value=None)       # 均质供应商编号 留空
    ws.cell(row=_INFO_ROW_CAT, column=6, value=st.color)
    ws.cell(row=_INFO_ROW_CAT, column=8, value=None)       # 回收料类型 留空

    # --- 成分行
    total = 0.0
    for i, c in enumerate(rec.components):
        r = st.tpl_comp_first + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=c.substance)
        # 2026-08-31 需求：无 CAS NO 的保密物质，CAS 栏填 00-00-00（而不是留空或写
        # Trade Secret）。判定为保密的两种情况都覆盖：CAS 为空 / CAS 是保密占位文字。
        ws.cell(row=r, column=3,
                value=SECRET_CAS_VALUE if c.secret else c.cas)
        ws.cell(row=r, column=4, value='是' if c.secret else '否')
        val = round(float(c.content), st.content_decimals)
        ws.cell(row=r, column=5, value=val)
        total += val
    if abs(total - 100.0) > 1e-6:
        warns.append(f'含量合计 {total:.4f}%，未精确等于 100%')

    # --- 结果行
    r_res = st.tpl_result_row + k_comp
    ws.cell(row=r_res, column=1, value='*测试结果')
    for item, col in RESULT_ITEMS:
        v = rec.results.get(item)
        cell = ws.cell(row=r_res, column=col, value=v if v else 'NA')
        if item in rec.inferred:
            cell.fill = RED_FILL
        elif item in rec.conflicts:
            cell.fill = RED_FILL

    # --- 报告行
    r_rep = st.tpl_report_first + k_comp
    for i, rp in enumerate(rec.reports):
        r = r_rep + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=rp.agency)
        cell_no = ws.cell(row=r, column=3, value=rp.report_no)
        if not rp.report_no:
            cell_no.fill = RED_FILL
            warns.append(f'第 {i + 1} 条报告缺失报告编号')
        dt = _to_dt(rp.report_date)
        if dt:
            cell_d = ws.cell(row=r, column=4, value=dt)
            cell_d.number_format = 'yyyy-mm-dd'
        else:
            cell_d = ws.cell(row=r, column=4, value=rp.report_date)
            cell_d.fill = RED_FILL

    # --- 隐藏「填写范例」页（保留，供 H2 下拉引用不被破坏）
    for name in wb.sheetnames[1:]:
        if '范例' in name or 'example' in name.lower():
            wb[name].sheet_state = 'hidden'

    wb.save(out_path)
    return warns


def _to_dt(s: str):
    if not s:
        return None
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d'):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    return None
