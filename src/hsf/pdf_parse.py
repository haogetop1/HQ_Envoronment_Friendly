# -*- coding: utf-8 -*-
"""检测报告 PDF 解析。

从一份 PDF 中提取：
  * 报告编号（Report No.）
  * 报告日期（Issue Date / Date）
  * 测试机构（SGS / 华测 / 美信 ...）
  * 29 项管控物质的测试结果

结果表解析采用「表头定位 + 列切片」：
  1. 找出表头行（含 Item 列与 Result 列，列数 >= 3），记录各列的字符区间；
  2. 以相邻列区间的中点作为切分边界，把数据行按同样的边界切片；
  3. 取 Result 列的值。
该方式不依赖单元格是否为空，也不受中英混排表头的干扰，可兼容：
  SGS 广分   Item | Limit | Unit | MDL | A7
  SGS 台湾   Item | Method | Unit | MDL | Result | Limit
  AWSC       Item | Result | MDL | (Limit)
  美信       Item | Unit | MDL | Test Results | Limit

版面模式（layout）用列切片；普通模式（plain）无列对齐，改用「自右向左按列序取值」。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .config import AGENCY_RULES, ITEM_LABELS
from .ocr import ocr_available, ocr_pdf_lines
from .util import fmt_date, parse_date, read_pdf_text

# ------------------------------------------------------------ 词法

_SPLIT_RE = re.compile(r'\S+(?:[ ]\S+)*')


def split_groups(line: str):
    """按 1 个空格切分成 [(start, end, text)]；2 格以上视为列间隔。"""
    return [(m.start(), m.end(), m.group(0)) for m in _SPLIT_RE.finditer(line)]


# ------------------------------------------------------------ 表头识别

_SKIP_WORDS = ('item', 'method', 'unit', 'mdl', 'cas', 'limit', 'sample',
               'no.', 'requirement', 'conclusion', 'remark', 'note')
_SAMPLE_RE = re.compile(r'^(?:[A-Z]{1,3}\d{1,3}|\d{3,4}|S\d*)$')


def _pick_result_col(texts) -> int:
    """在表头各组里挑出「结果列」的下标。"""
    for j, t in enumerate(texts):
        low = t.lower()
        if any(w in low for w in _SKIP_WORDS):
            continue
        if 'result' in low or 'measured' in low:
            return j
        if _SAMPLE_RE.match(t.strip()):
            return j
    return -1


@dataclass
class Table:
    cols: list                 # [(start, end, text)]
    bnds: list                 # 相邻列之间的切分中点
    result_idx: int
    has_limit: bool

    def slice(self, line: str) -> list:
        out = []
        n = len(self.cols)
        for j in range(n):
            lo = 0 if j == 0 else self.bnds[j - 1]
            hi = len(line) + 999 if j == n - 1 else self.bnds[j]
            out.append(line[int(lo):int(hi)].strip())
        return out


def _make_table(groups) -> Table | None:
    if len(groups) < 3:
        return None
    texts = [g[2] for g in groups]
    if not any('item' in t.lower() for t in texts):
        return None
    ridx = _pick_result_col(texts)
    if ridx < 0:
        return None
    bnds = [(groups[i][1] + groups[i + 1][0]) / 2.0 for i in range(len(groups) - 1)]
    has_limit = any('limit' in t.lower() for t in texts)
    return Table(groups, bnds, ridx, has_limit)


def _is_header(line: str):
    return _make_table(split_groups(line))


_STOP_RE = re.compile(
    r'^\s*(Note|Notes|Remark|Remarks|Conclusion|Appendix|Abbreviation|'
    r'\*+|=|\(1\)|\(2\)|This document|Unless otherwise|End of|樣品|备注|備註)', re.I)
_PAGE_RE = re.compile(r'Page\s+\d+\s+of\s+\d+|第\s*\d+\s*頁|第\s*\d+\s*页', re.I)


# ------------------------------------------------------------ 数值规范化

_ND_RE = re.compile(r'^[<＜]?\s*n\.?\s*d\.?$', re.I)
_NUM_RE = re.compile(r'^[<＞＞~]?\s*(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*(?:[▲*†‡]+)?$')
_UNIT_ONLY_RE = re.compile(r'^(?:mg/kg|mg\s*kg|ppm|ppb|%|µg/cm|ug/cm|mg/l|g/kg)$', re.I)


def canon_value(tok: str):
    """把单元格文本规范成 'ND' / 数字字符串；无法识别返回 None。"""
    if tok is None:
        return None
    t = tok.strip()
    t = t.replace('−', '-').replace('–', '-')
    t = re.sub(r'[*▲†‡]+$', '', t).strip()
    t2 = re.sub(r'\s+', '', t)
    if not t2 or t2 in ('-', '--', '—', '/', 'N/A', 'NA', 'n.a.'):
        return None
    if _ND_RE.match(t2):
        return 'ND'
    if _UNIT_ONLY_RE.match(t2):
        return None
    if t2.endswith('mg/kg') or t2.endswith('ppm') or t2.endswith('%'):
        # 「131 mg/kg」这类：数字本身才是结果
        m = re.match(r'^([\d.]+)', t2)
        if m:
            t, t2 = m.group(1), m.group(1)
    m = _NUM_RE.match(t) or _NUM_RE.match(t2)
    if m:
        raw = m.group(1).replace(',', '')
        try:
            v = float(raw)
        except ValueError:
            return None
        if t.lstrip().startswith('<'):
            return 'ND'
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        return ('%.6f' % v).rstrip('0').rstrip('.')
    return None


# ------------------------------------------------------------ 项目匹配

# 部分 PDF（如深圳美信）字符编码偏移后括号会变成这些符号
_PUNCT_MAP = {'˄': '(', '˅': ')', '（': '(', '）': ')', 'ǃ': '!', '˖': ':', 'ʼ': "'"}


def _norm_punct(s: str) -> str:
    for a, b in _PUNCT_MAP.items():
        s = s.replace(a, b)
    return s


def _match_item(text: str, max_start: int = 10 ** 6):
    """在文本中匹配管控物质，返回 (item, 命中结束位置)；未命中返回 (None, 0)。"""
    flat = re.sub(r'\s+', ' ', _norm_punct(text)).strip()
    if not flat:
        return None, 0
    best = (None, 0, 10 ** 6)
    for item, patterns in ITEM_LABELS.items():
        for p in patterns:
            m = re.search(p, flat, re.I | re.M)
            if m and m.start() <= max_start and m.start() < best[2]:
                best = (item, m.end(), m.start())
    return (best[0], best[1]) if best[0] else (None, 0)


def _item_of(text: str):
    return _match_item(text)[0]


# ------------------------------------------------------------ 版面模式 A：按 token 中心归列

def _assign(groups, cols):
    """把数据行的各组按「中心距离最近」归入表头列，返回 {列下标: 合并文本}。"""
    out = {}
    for g in groups:
        c = (g[0] + g[1]) / 2.0
        best, bd = -1, None
        for j, col in enumerate(cols):
            s, e = col[0], col[1]
            d = 0.0 if s - 3 <= c <= e + 3 else min(abs(c - s), abs(c - e))
            if bd is None or d < bd:
                best, bd = j, d
        if bd is not None and bd <= 60:
            out[best] = (out.get(best, '') + ' ' + g[2]).strip()
    return out


def _parse_block_tokens(lines, start, table: Table, results: dict, max_lines: int = 200):
    ridx = table.result_idx
    for k in range(start, min(start + max_lines, len(lines))):
        line = lines[k]
        if not line.strip() or _STOP_RE.match(line):
            break
        if _is_header(line):
            break
        cells = _assign(split_groups(line), table.cols)
        name = ' '.join(cells[j] for j in sorted(cells) if j < ridx)
        item = _item_of(name)
        if not item or item in results:
            continue
        val = canon_value(cells.get(ridx))
        if val is None:
            cands = [canon_value(cells[j]) for j in sorted(cells) if j != ridx]
            cands = [c for c in cands if c]
            val = cands[-1] if cands else None
        if val is not None:
            results[item] = val


def _parse_layout_tokens(lines, results: dict):
    for i, line in enumerate(lines):
        t = _is_header(line)
        if t:
            _parse_block_tokens(lines, i + 1, t, results)


# ------------------------------------------------------------ 版面模式 B：按列区间切片（兜底）

def _parse_block_layout(lines, start, table: Table, results: dict, max_lines: int = 200):
    ridx = table.result_idx
    for k in range(start, min(start + max_lines, len(lines))):
        line = lines[k]
        if not line.strip() or _STOP_RE.match(line):
            break
        if _is_header(line):
            break
        cells = table.slice(line)
        name = ' '.join(c for c in cells[:ridx] if c)
        item = _item_of(name)
        if not item:
            continue
        if item in results:
            continue
        val = canon_value(cells[ridx])
        if val is None:
            # 结果列被切碎（如 '13' | '1 mg/kg'）时，向右拼接
            joined = ''.join(cells[ridx:ridx + 2]).strip()
            val = canon_value(joined)
        if val is None and ridx > 0:
            val = canon_value(cells[ridx - 1] + cells[ridx])
        if val is None and ridx + 1 < len(cells):
            val = canon_value(cells[ridx + 1])
        if val is not None:
            results[item] = val


def _parse_layout_slice(lines, results: dict):
    for i, line in enumerate(lines):
        t = _is_header(line)
        if t:
            _parse_block_layout(lines, i + 1, t, results)


# ------------------------------------------------------------ 普通模式：自右向左按列序取值

_KEY_RE = re.compile(r'\b(item|result|mdl|limit|unit|method)\b', re.I)
_UNIT_TOKEN_RE = re.compile(r'^(mg/kg|mg\s*kg|ppm|ppb|%|mg/l|g/kg|µg/cm\S*)$', re.I)


def _plain_hint(text: str) -> str:
    """判断结果列位置：'last2'（Result 之后还有 Limit 列）或 'last'。"""
    low = text.lower()
    ir = low.find('result')
    il = low.find('limit')
    if ir >= 0 and il > ir:
        return 'last2'
    return 'last'


def _parse_plain(lines, results: dict):
    """普通模式兜底：标签在行首 + 短行，按表头提示取结果列。

    表头可能被拆成多行（美信即是），故用滑动缓冲区把连续的短表头碎片拼起来。
    """
    hint = 'last'
    buf = ''
    for line in lines:
        s = line.strip()
        if not s:
            buf = ''
            continue
        if _PAGE_RE.search(s) or len(s) > 110:
            buf = ''
            continue
        if _KEY_RE.search(s):
            buf = (buf + ' ' + s).strip()
            if 'item' in buf.lower() and re.search(r'\b(mdl|result|unit|limit)\b', buf, re.I):
                hint = _plain_hint(buf)
            continue
        if _STOP_RE.match(s):
            buf = ''
            continue
        buf = ''
        item, end = _match_item(s, max_start=45)
        if not item or item in results:
            continue
        toks = s[end:].split()
        vals = [canon_value(t) for t in toks if not _UNIT_TOKEN_RE.match(t)]
        vals = [v for v in vals if v]
        if not vals:
            continue
        val = vals[-2] if (hint == 'last2' and len(vals) >= 2) else vals[-1]
        results[item] = val


# ------------------------------------------------------------ 报告编号 / 日期 / 机构

_NO_PATTERNS = [
    r'(?:Test\s*)?(?:Report|Raport)\s*(?:No|No\.|Number|#)\s*[:.：]?\s*'
    r'([A-Za-z0-9][A-Za-z0-9\-_/]{6,40})',
    r'(?:報告編號|报告编号)\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-_/]{6,40})',
    r'\bNo\.\s*[:：]?\s*([A-Z]{2,6}[0-9][A-Za-z0-9\-_/]{5,34})',
]
_DATE_PATTERNS = [
    r'(?:Issue\s*Date|Report\s*Date|Date\s*of\s*Issue|Issued\s*Date|'
    r'報告日期|报告日期)\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9,\.\-/ ]{0,24})?',
    r'\bDate\s*[:：]\s*([A-Za-z0-9][A-Za-z0-9,\.\-/ ]{0,24})?',
]
_BAD_NO = re.compile(r'^(page|tel|fax|www|http|email|报告|report|cust|addre)', re.I)
_JOB_RE = re.compile(r'(job|工作|委托)\s*no', re.I)


_DATEISH_RE = re.compile(r'^[\dA-Za-z,\.\-/ ]{4,24}$')


def _next_texts(lines, i, span=4):
    """处理「标签一行、值一行」的排版：把后续短小且像日期的行接上来。"""
    out = [lines[i]]
    buf = lines[i]
    got = 0
    for k in range(i + 1, min(i + 1 + span, len(lines))):
        nxt = lines[k].strip()
        if not nxt or not _DATEISH_RE.match(nxt):
            continue
        buf = buf.rstrip() + ' ' + nxt
        out.append(buf)
        got += 1
        if got >= 2:
            break
    return out


def _find_report_no(*line_lists) -> str:
    for pattern in _NO_PATTERNS:
        for lines in line_lists:
            for i, line in enumerate(lines[:120]):
                if _JOB_RE.search(line):
                    continue
                for m in re.finditer(pattern, line, re.I):
                    cand = re.sub(r'\s+', '', m.group(1)).strip('.,;:')
                    if _BAD_NO.match(cand) or len(cand) < 7 or not re.search(r'\d', cand):
                        continue
                    return cand
    return ''


def _find_date(*line_lists) -> str:
    for pattern in _DATE_PATTERNS:
        for lines in line_lists:
            for i, line in enumerate(lines[:120]):
                m = re.search(pattern, line, re.I)
                if not m:
                    continue
                for cand_text in _next_texts(lines, i):
                    m2 = re.search(pattern, cand_text, re.I)
                    if not m2 or not (m2.group(1) or '').strip():
                        continue
                    d = parse_date(m2.group(1))
                    if d:
                        return fmt_date(d)
    for lines in line_lists:
        for line in lines[:40]:
            for m in re.finditer(r'(\d{1,2}[-\s.][A-Za-z]{3,9}[-\s.]\d{4}|'
                                 r'[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}|'
                                 r'\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})', line):
                d = parse_date(m.group(1))
                if d:
                    return fmt_date(d)
    return ''


def _find_agency(*line_lists, file_name: str = '') -> str:
    parts = ['\n'.join(l[:400] for l in lines[:200]) for lines in line_lists]
    hay = '\n'.join(parts) + '\n' + (file_name or '')
    for keys, short in AGENCY_RULES:
        for k in keys:
            if k.lower() in hay.lower():
                return short
    return ''


# ------------------------------------------------------------ 对外接口

@dataclass
class ReportInfo:
    source: str = ''                                   # 内嵌时的原始文件名
    report_no: str = ''
    report_date: str = ''
    agency: str = ''
    results: dict = field(default_factory=dict)         # item -> 'ND' / '217'
    warnings: list = field(default_factory=list)
    ok: bool = True

    @property
    def label(self) -> str:
        return self.report_no or self.source


def parse_pdf(data: bytes, source: str = '', file_name_hint: str = '') -> ReportInfo:
    """解析一份检测报告 PDF。失败时返回 ok=False。"""
    txt = read_pdf_text(data)
    if txt is None:
        info = ReportInfo(source=source, ok=False)
        info.warnings.append('PDF 无法打开或已损坏')
        return info

    info = ReportInfo(source=source)
    results: dict = {}
    _parse_layout_tokens(txt.layout_lines, results)     # 主：按 token 中心归列
    _parse_layout_slice(txt.layout_lines, results)      # 补：按列区间切片
    _parse_plain(txt.plain_lines, results)              # 兜底：普通模式补缺

    info.results = results
    info.report_no = _find_report_no(txt.layout_lines, txt.plain_lines)
    info.report_date = _find_date(txt.layout_lines, txt.plain_lines)
    info.agency = _find_agency(txt.layout_lines, txt.plain_lines,
                               file_name=file_name_hint or source)

    # 扫描件兜底（2026-08-31）：图片型 PDF 抽不出文本，改用 OCR。
    # 只在「编号或结果没抓到」时启用——OCR 每页约 6 秒，不能对正常 PDF 乱用。
    if (not info.report_no or not results):
        if ocr_available():
            ocr_lines = ocr_pdf_lines(data)
            if ocr_lines:
                ocr_res: dict = {}
                _parse_layout_tokens(ocr_lines, ocr_res)
                _parse_layout_slice(ocr_lines, ocr_res)
                _parse_plain(ocr_lines, ocr_res)
                for k, v in ocr_res.items():
                    results.setdefault(k, v)      # 已有的结果不被 OCR 覆盖
                info.results = results
                if not info.report_no:
                    info.report_no = _find_report_no(ocr_lines, ocr_lines)
                if not info.report_date:
                    info.report_date = _find_date(ocr_lines, ocr_lines)
                if not info.agency:
                    info.agency = _find_agency(ocr_lines, ocr_lines,
                                               file_name=file_name_hint or source)
                info.warnings.append(
                    f'文本层抽取不到内容，已改用 OCR 识别（{len(ocr_lines)} 行）')
        else:
            # 精简版：该 OCR 却没 OCR。不留痕的话，用户只看到结果变少，
            # 无从判断是「报告本就没测」还是「工具没识别出」。
            info.warnings.append(
                f'{source}：本需 OCR 才能识别，但当前程序未集成 OCR（精简版），'
                '编号将填 TBD00000000、未测项填 NA，需人工补录')

    if not info.report_no:
        info.warnings.append('未识别到报告编号')
    if not info.report_date:
        info.warnings.append('未识别到报告日期')
    if not info.agency:
        info.warnings.append('未识别到测试机构')
    if not results:
        info.warnings.append('未解析到任何测试结果')
    return info


# ------------------------------------------------------------ 多份报告合并

def _rank(v: str) -> float:
    """排序权重：数字 > ND；数字取数值大小。"""
    if v == 'ND':
        return -1.0
    try:
        return float(v)
    except ValueError:
        return -2.0


def merge_results(infos: list[ReportInfo]):
    """多份报告逐项取最不利值（数字大于 ND，数字取最大值）。

    返回 (结果字典, 冲突项集合)。
    """
    merged: dict = {}
    conflict = set()
    for info in infos:
        for item, v in info.results.items():
            if item not in merged:
                merged[item] = v
                continue
            a, b = merged[item], v
            if a == b:
                continue
            conflict.add(item)
            merged[item] = a if _rank(a) >= _rank(b) else b
    return merged, conflict
