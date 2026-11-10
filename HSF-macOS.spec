# Build on an Apple Silicon Mac with Python 3.12.
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files('tkinterdnd2') + collect_data_files('numbers_parser')
template = Path('assets/均质创建导入模板.xlsx')
if template.is_file():
    datas.append((str(template), 'assets'))
a = Analysis(['run_gui.py'], pathex=['src'], binaries=[], datas=datas,
    hiddenimports=['run_cli'],
    excludes=['pytest','reportlab','xlwt','matplotlib','pandas','IPython','numpy'],
    hookspath=[], runtime_hooks=[], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='HSFHelper',
    debug=False, strip=False, upx=False, console=False, target_arch='arm64')
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='HSFHelper')
app = BUNDLE(coll, name='HSF均质表助手.app', bundle_identifier='local.hsf.macassistant',
    info_plist={'CFBundleName':'HSF均质表助手', 'CFBundleDisplayName':'HSF均质表助手',
                'CFBundleShortVersionString':'1.1.4', 'CFBundleVersion':'7',
                'NSHighResolutionCapable':True, 'LSMinimumSystemVersion':'14.0'})
