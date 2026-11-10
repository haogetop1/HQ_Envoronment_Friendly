# HSF 均质表助手（macOS）

双击 Apple Silicon 版「HSF均质表助手.app」。选择带有内嵌 PDF 的材料成分表和均质创建导入模板（`.xlsx`）。确认芯片编号和保存位置后点击「开始生成」。生成完成后可在 Finder 打开本次输出文件夹。

应用在本机处理 `.xls` 和 `.xlsx`，从内嵌 PDF 的文字层提取报告数据。纯图片扫描件会保留为 PDF 并提示人工复核。会生成各材料的均质导入表、HSF PDF、MSDS PDF、汇总表和日志。支持原项目 K/M 列布局，也能从完整表头识别 J/L 列布局。加密的工作簿需要在有权限的原环境中导出可读取的副本，保留内嵌附件。

Numbers（`.numbers`，也接受 `.number` 文件名）会先把材料成分表另存为 `原文件名-转换.xlsx`，保留单元格值、合并区域和原行列位置。公式按当前值导出，浮动图标不写入转换表。只需转 Excel 时，清空模板输入框即可，也可运行 `.venv/bin/python run_cli.py 源表.numbers 输出目录`。

选择导入模板后，程序重新打开转换出的 Excel，以其中的材料成分和重量生成均质导入表。原 Numbers 文件的图标坐标用于定位附件，并从「Numbers 原始报告文件夹」按文件名匹配原始 PDF，忽略图标换行、空格和大小写。同名候选有多份时会标记待核对。`Numbers附件核对.xlsx` 记录材料成分和缺件情况。检测报告齐全时再生成均质导入 `.xlsx`。命令行可用 `--pdf-dir` 指定报告目录。

填写沿用原项目的默认值（SAC、Metal、TBD、黑色、Ⅰ级）和测试结果合并规则。查看汇总表和红色单元格，对照原报告复核补填的 ND、默认分类和附件材料型号，再导入业务系统。

## 运行或重新打包

在 Apple Silicon Mac 上使用带 Tk 的 Python 3.12，从此目录执行。

```sh
python3.12 -m venv .venv
.venv/bin/python -m pip install -r requirements-macos-lock.txt
.venv/bin/python run_gui.py
.venv/bin/python -m pytest tests -q
.venv/bin/python -m PyInstaller --noconfirm HSF-macOS.spec
```

打包产物位于 `dist/HSF均质表助手.app`，使用本地签名。公开分支仅提供源码和合成测试数据，运行时请选择自己的导入模板。如果需要在本地打包时预选模板，可在构建前将模板放到 `assets/均质创建导入模板.xlsx`。命令行也可通过 `.venv/bin/python run_cli.py 源表.xls 模板.xlsx 输出目录` 运行。

这是 macOS 分支，沿用原项目的许可声明，保留所有权利。Windows 版见 [main 分支](https://github.com/haogetop1/HQ_Envoronment_Friendly/tree/main)。

旧版 Excel 对象关联按 [Microsoft MS-XLS 的 Embedding Storage 规范](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b406ade0-fb1c-4512-bff2-b576fdfff545)读取工作表对象指定的存储编号。
