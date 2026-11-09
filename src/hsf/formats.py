"""Identify the container, never trust the filename extension."""
from pathlib import Path
from zipfile import is_zipfile

TSD_MESSAGE = ('文件经过企业加密（TSD），无法直接读取。请在有权限的原环境中解密导出，'
               '再选择未加密的 Excel 文件，并保留内嵌 PDF；修改扩展名不能解密。')


def workbook_format(path: str) -> str:
    with open(path, 'rb') as f:
        header = f.read(32)
    if header.startswith(b'%TSD-Header'):
        raise ValueError(TSD_MESSAGE)
    if header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
        import olefile
        with olefile.OleFileIO(path) as ole:
            if ole.exists('EncryptedPackage'):
                raise ValueError('Excel 文件有打开密码，请先用 Excel 解密另存为，再重试。')
        return 'xls'
    if is_zipfile(path):
        return 'xlsx'
    raise ValueError(f'{Path(path).name} 不是可读取的 Excel 文件，可能已加密或损坏。')
