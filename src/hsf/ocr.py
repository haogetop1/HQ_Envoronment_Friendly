# -*- coding: utf-8 -*-
"""扫描件 OCR 兜底。

纯图片型 PDF（纸质盖章件扫描，实测某份 25 页），pypdf 只能抽出零散字符，
报告编号和测试结果全都抓不到。这类 PDF 必须先光栅化、OCR、再把识别框按 y
坐标重建成文本行，才能喂给 pdf_parse 里那套已经调好的解析逻辑。

**触发条件刻意不用「文本密度」**——实测扫描件 532 字符/页，而正常的检测报告
也有 854 字符/页，阈值根本区分不开，会把好 PDF 也送去 OCR。改为
**「常规解析没抓到编号或测试结果」才启用**，精准且不给正常 PDF 白白加几十秒。

OCR 依赖是可选的：没装 rapidocr / pypdfium2 就静默降级，主流程照常跑。
"""
from __future__ import annotations

_ENGINE = None
_AVAILABLE: bool | None = None

# 最多 OCR 前几页。报告编号/日期在首页，测试结果一般在第 2-3 页。
# OCR 每页约 6 秒，不设上限的话 25 页的报告要跑两分多钟。
MAX_PAGES = 6
SCALE = 1.5


def ocr_available() -> bool:
    """OCR 依赖是否已安装（结果缓存，只探测一次）。"""
    global _AVAILABLE
    if _AVAILABLE is None:
        try:
            import pypdfium2  # noqa: F401
            from rapidocr_onnxruntime import RapidOCR  # noqa: F401
            _AVAILABLE = True
        except Exception:
            _AVAILABLE = False
    return _AVAILABLE


def _engine():
    global _ENGINE
    if _ENGINE is None and ocr_available():
        from rapidocr_onnxruntime import RapidOCR
        _ENGINE = RapidOCR()
    return _ENGINE


def boxes_to_lines(res, y_tol: float = 12.0) -> list:
    """把 OCR 的 [(box, text, score)] 按 y 中心聚成行，行内按 x 排序。

    OCR 返回的是一个个独立的文本框，直接拼接会把「项目名 / 方法 / 单位 / 结果」
    挤成乱序。按 y 归行、行内按 x 排序，才还原出表格的横向结构。

    行基准 y 用**滑动平均**而不是首框 y：扫描件常有轻微倾斜，同一行的框 y 会
    逐格下移，用首框做基准会把一行切成两半。
    """
    items = []
    for box, text, _score in (res or []):
        if not text:
            continue
        ys = [p[1] for p in box]
        xs = [p[0] for p in box]
        items.append((sum(ys) / len(ys), min(xs), text))
    items.sort()

    lines, cur, cy = [], [], None
    for y, x, text in items:
        if cy is None or abs(y - cy) <= y_tol:
            cur.append((x, text))
            cy = y if cy is None else (cy * (len(cur) - 1) + y) / len(cur)
        else:
            lines.append(' '.join(t for _, t in sorted(cur)))
            cur, cy = [(x, text)], y
    if cur:
        lines.append(' '.join(t for _, t in sorted(cur)))
    return lines


def ocr_pdf_lines(data: bytes, max_pages: int = MAX_PAGES,
                  scale: float = SCALE) -> list:
    """光栅化 + OCR + 重建行。返回文本行列表；依赖缺失或失败时返回 []。"""
    eng = _engine()
    if eng is None:
        return []
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(data)
    except Exception:
        return []

    lines: list = []
    for i in range(min(len(doc), max_pages)):
        try:
            img = doc[i].render(scale=scale).to_numpy()
            res, _ = eng(img)
        except Exception:
            continue
        lines.extend(boxes_to_lines(res))
    return lines
