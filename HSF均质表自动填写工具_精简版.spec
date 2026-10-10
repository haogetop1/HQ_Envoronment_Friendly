# -*- mode: python ; coding: utf-8 -*-
"""精简版打包配置：去掉 OCR 依赖，体积从完整版的 111 MB 降到 35.1 MB。

2026-08-31 实测：去掉 OCR 后仍有 35.1 MB（不是预想的 15 MB），剩余大头是
cryptography + numpy + tkinterdnd2。OCR 栈本身约 76 MB。

**降级范围（实测，比原先以为的要广）**：
不只是「扫描件（图片型 PDF）」会退化。OCR 的触发条件是「常规三路解析没抓到
编号或测试结果」，所以凡是走到 OCR 兜底的报告都会退化，包括：

- 扫描件（无文本层，纯图片型 PDF）
- 有文本层、但版面解析失败的报告 —— 如 RR12346 的 CopperFoil / NiPlating / PdPlating

退化表现：报告编号填 TBD00000000、测试结果未测项填 NA。
Pb/Cd/Hg/Cr6+ 这 4 个必填项仍由领域规则补 ND，**不会出现 NA**。

RR12346 实测 25 份表里有 **4 份**受影响（Ceramic / CopperFoil / NiPlating /
PdPlating），共 57 条非必填项由具体值或 ND 降级为 NA；
RR12345 实测 **0 份**受影响，与完整版逐格一致。

OCR 依赖是可选依赖（见 src/hsf/ocr.py）——`ocr_available()` 探测不到
rapidocr / pypdfium2 就返回 False 并降级，主流程照常跑。
所以这里只要不把它们打进去即可，主流程代码一行都不用改。

**降级必须可见，不能静默**（2026-08-31 补）：`pipeline.run()` 启动时若检测到
OCR 不可用会打 WARN，`pdf_parse` 对每份「本该 OCR 却没 OCR」的报告也会留告警。
否则用户只看到数据变少，分不清是「报告本就没测」还是「工具没识别出」。

什么时候用哪个：
  源表里的报告都是文本型 PDF  ->  精简版（启动快、好分发）
  有扫描件、想自动识别          ->  完整版（HSF均质表自动填写工具.exe）
"""
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['hsf.config', 'hsf.util', 'hsf.ole_extract', 'hsf.xls_source', 'hsf.pdf_parse', 'hsf.pdfout', 'hsf.xlsx_build', 'hsf.pipeline', 'hsf.gui', 'hsf.ocr']
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# pypdf 解析 AES 加密的 PDF 需要 cryptography。
# 2026-08-31：之前没打包它，遇到加密报告会抛 DependencyError，
# 异常一路冒泡导致整个材料块的 xlsx 都不生成（RR12346 的 AuPlating 就这么丢的）。
# 精简版同样必须带上——这条不是可选依赖。
tmp_ret = collect_all('cryptography')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# ==== 以下是完整版有、精简版刻意去掉的部分 ====
# tmp_ret = collect_all('rapidocr_onnxruntime')   # 含 ONNX 模型，约 40 MB
# tmp_ret = collect_all('pypdfium2')              # 光栅化，约 5 MB
# 连带还会甩掉 onnxruntime / opencv-python / shapely / numpy 的大部分体积。

# 显式排除，防止被 hiddenimport 或依赖链顺带拖进来
excludes = [
    'rapidocr_onnxruntime',
    'pypdfium2',
    'onnxruntime',
    'cv2',
    'shapely',
    'PyQt5',
    'matplotlib',
    'scipy',
    'pandas',
]


a = Analysis(
    ['run_gui.py'],
    pathex=['src'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='HSF均质表自动填写工具_精简版',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
