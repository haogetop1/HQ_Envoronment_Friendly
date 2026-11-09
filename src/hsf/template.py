"""Validate the supplied import schema and preserve Excel-only dropdowns."""
import copy
import os
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from .config import RESULT_ITEMS
from .formats import workbook_format

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def validate_template(path):
    if workbook_format(path) != 'xlsx':
        raise ValueError('模板须为真正的 .xlsx 文件。')
    with open(path,'rb') as f:
        w = openpyxl.load_workbook(f,read_only=True,data_only=True)
        try:
            s = w.worksheets[0]
            expected = {'A2':'*均质名称','A5':'*序号','B5':'*物质名称',
                        'A12':'该行不允许删除','A15':'*测试结果',
                        'B17':'*测试机构','C17':'*测试报告编号','A22':'该行不允许删除'}
            if any(s[cell].value!=value for cell,value in expected.items()):
                raise ValueError('模板结构与均质创建导入模板不一致，请使用随附模板。')
            if any(s.cell(14,col).value!=item for item,col in RESULT_ITEMS):
                raise ValueError('模板测试项目顺序不匹配，已停止以避免填错列。')
        finally:
            w.close()


def expand_validations(ws, first, marker, count):
    if not count:
        return
    for rule in ws.data_validations.dataValidation:
        updated = MultiCellRange()
        for ref in rule.ranges:
            r = CellRange(str(ref))
            if r.min_row>=marker:
                r.shift(row_shift=count)
            elif r.max_row>=marker or (r.min_row>=first and r.max_row==marker-1):
                r.max_row += count
            updated.add(r)
        rule.sqref = updated


def restore_extensions(template, output, component_extra=0, report_extra=0):
    # openpyxl drops x14 validation. Preserve those elements from the original ZIP.
    extensions = {}
    with ZipFile(template) as z:
        for name in z.namelist():
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                root = ET.fromstring(z.read(name))
                ext = root.find(f'{{{S}}}extLst')
                if ext is not None:
                    ext = copy.deepcopy(ext)
                    if name=='xl/worksheets/sheet1.xml':
                        for node in ext.iter():
                            if node.tag.endswith('}sqref') and node.text:
                                refs = []
                                for value in node.text.split():
                                    r=CellRange(value)
                                    for marker,count in [(12,component_extra),(22+component_extra,report_extra)]:
                                        if count and r.min_row>=marker: r.shift(row_shift=count)
                                        elif count and r.max_row>=marker: r.max_row+=count
                                    refs.append(str(r))
                                node.text=' '.join(refs)
                    extensions[name]=ext
    if not extensions:
        return
    temp = str(output)+'.tmp'
    try:
        with ZipFile(output) as src, ZipFile(temp,'w',ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename in extensions:
                    root = ET.fromstring(data)
                    root.append(extensions[info.filename])
                    data = ET.tostring(root,encoding='utf-8',xml_declaration=True)
                dst.writestr(info,data)
        os.replace(temp,output)
    finally:
        Path(temp).unlink(missing_ok=True)
