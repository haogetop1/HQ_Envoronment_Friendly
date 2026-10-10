# -*- coding: utf-8 -*-
"""通用工具：文件名清洗、日期解析、PDF 文本抽取与乱码修复。"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime

from pypdf import PdfReader

# ---------------------------------------------------------------- 文件名 / 文本

_INVALID = r'[\\/:*?"<>|\r\n\t]'


def safe_name(s: str, default: str = 'noname') -> str:
    """清洗成合法的 Windows 文件名片段。"""
    s = unicodedata.normalize('NFKC', str(s or ''))
    s = re.sub(_INVALID, '_', s)
    s = ''.join(ch for ch in s if ord(ch) >= 32)
    s = s.strip().strip('.')
    return s or default


def dedot(s: str) -> str:
    return s


def remove_spaces(s: str) -> str:
    """材料名称：删除所有空白字符。"""
    s = unicodedata.normalize('NFKC', str(s or ''))
    return re.sub(r'\s+', '', s)


# 中英文括号及其内容，如 Die1(EMMJR68) -> Die1、标题（备注） -> 标题
_PAREN_RE = re.compile(r'[（(][^（）()]*[）)]')


def strip_paren(s: str) -> str:
    """去掉括号及其内部内容（2026-08-31 需求：均质名称不取括号内的型号）。

    先 NFKC 归一，让全角括号（（））与半角（()）一并处理。
    括号内容通常是厂家内部型号（EMMJR68 / PHC751C），不该进均质名称。
    """
    s = unicodedata.normalize('NFKC', str(s or ''))
    return _PAREN_RE.sub('', s)


def material_display(s: str) -> str:
    """材料显示名：去括号内容 + 去所有空白。Die1(EMMJR68) -> Die1。"""
    return remove_spaces(strip_paren(s))


# 2026-08-31 需求：CAS No 只允许数字和 '-'，其余字符一律剔除。
# 源表里存在 7440-57-5(金) 这类混了中文注释的写法，导入系统会判为非法值。
_CAS_STRIP_RE = re.compile(r'[^0-9\-]')


def clean_cas(s: str) -> str:
    """CAS No 清洗：只保留数字和 '-'。

        '7440-57-5(金)'  -> '7440-57-5'
        'Trade secret'   -> ''      （清洗后为空 -> 判定为保密成分）
        '147 - 14 - 8'   -> '147-14-8'

    先 NFKC 归一，全角数字／全角横线一并转半角。
    """
    s = unicodedata.normalize('NFKC', str(s or ''))
    return _CAS_STRIP_RE.sub('', s)


def extract_prefix(source_path: str) -> str | None:
    """从源文件名中提取芯片编号，如「材料成分表 RR12345.xls」-> RR12345。

    编号形如 `RR` + 3~8 位数字，前缀与数字之间的空格／横线会被一并吃掉。

    抓不到时返回 None（2026-08-31：不再硬编码兜底，由调用方弹窗问用户）。
    """
    name = source_path.rsplit('\\', 1)[-1].rsplit('/', 1)[-1]
    m = re.search(r'\b(RR\s*-?\s*\d{3,8})\b', name, re.I)
    if m:
        return re.sub(r'[\s-]', '', m.group(1)).upper()
    m = re.search(r'([A-Z]{1,4}-?\d{4,8})', name)
    if m:
        return m.group(1).upper()
    return None


# ---------------------------------------------------------------- PDF 文本

_SCORE_WORDS = ['Report', 'Test', 'Date', 'Page', 'mg/kg', 'MDL', 'Sample', 'Result',
                'Item', 'ND', 'Limit', 'Unit', 'Method', 'Customer', 'Issue',
                'RoHS', 'Halogen', 'Appendix', 'Conclusion']


def _shift(t: str, s: int) -> str:
    out = []
    for ch in t:
        o = ord(ch)
        if ch in ' \n\r\t' or o > 0xFF:
            out.append(ch)
        else:
            out.append(chr((o + s) & 0xFF))
    return ''.join(out)


def _score(t: str) -> int:
    return sum(t.count(w) for w in _SCORE_WORDS)


def repair_encoding(text: str):
    """修复字符编码整体偏移的 PDF 文本（如深圳美信报告偏移 +29）。

    返回 (修复后文本, 使用的偏移量)。偏移量 0 表示无需修复。
    """
    base = _score(text)
    if base >= 8:                      # 文本本身已经可读
        return text, 0
    best_s, best = 0, base
    for s in range(1, 256):
        sc = _score(_shift(text, s))
        if sc > best:
            best_s, best = s, sc
    if best < 8 or best < base * 2:
        return text, 0
    return _shift(text, best_s), best_s


class PdfText:
    """一份 PDF 的两种文本视图：普通模式 / 版面模式（均已做编码修复）。"""

    def __init__(self, raw_plain: str, raw_layout: str):
        self.plain, self.shift_plain = repair_encoding(raw_plain)
        self.layout, self.shift_layout = repair_encoding(raw_layout)
        self.shift = self.shift_plain or self.shift_layout

    @property
    def layout_lines(self) -> list:
        return self.layout.split('\n')

    @property
    def plain_lines(self) -> list:
        return self.plain.split('\n')

    @property
    def all_lines(self):
        seen = set()
        for src in (self.layout, self.plain):
            for ln in src.split('\n'):
                k = ln.strip()
                if k and k not in seen:
                    seen.add(k)
                    yield ln


def read_pdf_text(data: bytes):
    """bytes -> PdfText；失败返回 None。"""
    import io
    try:
        reader = PdfReader(io.BytesIO(data))
    except Exception:
        return None
    plain, layout = [], []
    for page in reader.pages:
        try:
            plain.append(page.extract_text() or '')
        except Exception:
            plain.append('')
        try:
            layout.append(page.extract_text(extraction_mode='layout') or '')
        except Exception:
            layout.append('')
    return PdfText('\n'.join(plain), '\n'.join(layout))


# ---------------------------------------------------------------- 日期

_MONTHS = {m: i + 1 for i, m in enumerate(
    ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'])}
_MONTHS.update({'sept': 9})


def parse_date(text: str):
    """把各种写法的日期解析成 date 对象。"""
    if not text:
        return None
    t = text.strip().strip('.')
    t = re.sub(r'\s+', ' ', t)

    # 2026.04.13 / 2026/4/13 / 2026-04-13
    m = re.match(r'^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$', t)
    if m:
        y, mo, d = map(int, m.groups())
        return _mk(y, mo, d)

    # 13-Apr-2026 / 08-Dec-2025
    m = re.match(r'^(\d{1,2})[-\s.]([A-Za-z]{3,9})[-\s.](\d{4})$', t)
    if m:
        d, mon, y = m.group(1), m.group(2).lower(), m.group(3)
        if mon in _MONTHS:
            return _mk(int(y), _MONTHS[mon], int(d))

    # Sep 03, 2025 / Apr 14, 2026 / Jan 18, 2024
    m = re.match(r'^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})$', t)
    if m:
        mon, d, y = m.group(1).lower(), m.group(2), m.group(3)
        if mon in _MONTHS:
            return _mk(int(y), _MONTHS[mon], int(d))

    # 2025年8月1日 / 2025/8/1
    m = re.match(r'^(\d{4})\s*[年.\-/]\s*(\d{1,2})\s*[月.\-/]\s*(\d{1,2})', t)
    if m:
        y, mo, d = map(int, m.groups())
        return _mk(y, mo, d)
    return None


def _mk(y, m, d):
    try:
        return datetime(y, m, d).date()
    except Exception:
        return None


def fmt_date(d) -> str:
    return d.strftime('%Y-%m-%d') if d else ''
