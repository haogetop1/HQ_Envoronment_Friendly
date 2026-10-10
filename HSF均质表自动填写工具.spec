# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['hsf.config', 'hsf.util', 'hsf.ole_extract', 'hsf.xls_source', 'hsf.pdf_parse', 'hsf.pdfout', 'hsf.xlsx_build', 'hsf.pipeline', 'hsf.gui', 'hsf.ocr']
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# pypdf 解析 AES 加密的 PDF 需要 cryptography。
# 2026-08-31：之前没打包它，遇到加密报告会抛 DependencyError，
# 异常一路冒泡导致整个材料块的 xlsx 都不生成（RR12346 的 AuPlating 就这么丢的），
# 而且日志里只有一串 traceback，看不出是哪份报告出的问题。
tmp_ret = collect_all('cryptography')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 扫描件 OCR 兜底（可选依赖，缺了会静默降级）
# rapidocr_onnxruntime 自带 ONNX 模型，必须用 collect_all 才能把模型文件带上
tmp_ret = collect_all('rapidocr_onnxruntime')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pypdfium2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['run_gui.py'],
    pathex=['src'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='HSF均质表自动填写工具',
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
