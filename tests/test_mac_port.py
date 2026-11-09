import io
import struct
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

import openpyxl
import pytest
from reportlab.pdfgen.canvas import Canvas

from hsf.config import Settings, RESULT_ITEMS
from hsf.ole_extract import _parse_ole10native, extract_embedded_objects
from hsf.xls_source import SourceSheet, Component, normalize_content
from hsf.pipeline import Pipeline, _dedup
from hsf.xlsx_build import ComponentRow, MaterialRecord, ReportRow, build_workbook

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'


@pytest.fixture
def template(tmp_path):
    path = tmp_path / 'test-template.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for cell, value in {'A2':'*均质名称', 'A5':'*序号', 'B5':'*物质名称',
                        'A12':'该行不允许删除', 'A13':'HSF测试结果', 'A15':'*测试结果',
                        'B17':'*测试机构', 'C17':'*测试报告编号', 'A22':'该行不允许删除'}.items():
        sheet[cell] = value
    for item, column in RESULT_ITEMS:
        sheet.cell(14, column, item)
    sheet.merge_cells('A13:AC13')
    validation = openpyxl.worksheet.datavalidation.DataValidation(type='list', formula1='"是,否"')
    validation.add('D6:D11')
    sheet.add_data_validation(validation)
    workbook.save(path)

    # Excel also writes some dropdowns in the x14 extension rather than the standard list.
    x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
    xm = 'http://schemas.microsoft.com/office/excel/2006/main'
    key = 'xl/worksheets/sheet1.xml'
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts[key])
    extension = ET.SubElement(ET.SubElement(root, f'{{{S}}}extLst'), f'{{{S}}}ext',
                              {'uri': '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}'})
    validations = ET.SubElement(extension, f'{{{x14}}}dataValidations', {'count': '1'})
    rule = ET.SubElement(validations, f'{{{x14}}}dataValidation', {'type': 'list'})
    formula = ET.SubElement(rule, f'{{{x14}}}formula1')
    ET.SubElement(formula, f'{{{xm}}}f').text = '"选项一,选项二"'
    ET.SubElement(rule, f'{{{xm}}}sqref').text = 'H2'
    parts[key] = ET.tostring(root)
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for name, contents in parts.items():
            archive.writestr(name, contents)
    return path


def pdf_bytes(report='CANEC2612345678', pb='ND'):
    stream = io.BytesIO()
    c = Canvas(stream)
    for y, line in enumerate(['SGS Test Report', f'Report No.: {report}', 'Date: 2026-10-01',
                               f'Lead (Pb) {pb}', 'Cadmium (Cd) ND', 'Mercury (Hg) ND',
                               'Hexavalent Chromium (Cr6+) ND']):
        c.drawString(60, 780-y*24, line)
    c.showPage(); c.save()
    return stream.getvalue()


def source_fixture(path, *, missing_anchor=False, external=False, other_sheet=False):
    w = openpyxl.Workbook()
    if other_sheet:
        w.active.title = '说明'
        s = w.create_sheet('材料成分表')
        sheet_no = 2
    else:
        s = w.active; s.title = '材料成分表'; sheet_no = 1
    for row, sub, cas, mass in [(8, 'Copper', '7440-50-8', 1),
                                 (9, 'Tin', '7440-31-5', 3)]:
        for col, value in {3:'Supplier', 4:'Leadframe(model)', 6:sub, 7:cas, 8:mass}.items():
            s.cell(row, col, value)
    s.merge_cells('K8:K9'); s.merge_cells('M8:M9')
    s['A10'] = '合计'
    w.save(path)
    with ZipFile(path) as z: files = {n:z.read(n) for n in z.namelist()}
    key = f'xl/worksheets/sheet{sheet_no}.xml'
    root = ET.fromstring(files[key]); objects = ET.SubElement(root, f'{{{S}}}oleObjects')
    rels = ET.Element(f'{{{P}}}Relationships')
    for i, col in [(1,10), (2,12)]:
        obj = ET.SubElement(objects, f'{{{S}}}oleObject', {'progId':'AcroExch.Document.DC', 'shapeId':str(i), f'{{{R}}}id':f'rId{i}'})
        if not missing_anchor:
            props = ET.SubElement(obj, f'{{{S}}}objectPr')
            anchor = ET.SubElement(props, f'{{{S}}}anchor')
            for tag, row in [('from',7),('to',8)]:
                edge = ET.SubElement(anchor, f'{{{S}}}{tag}')
                for name, val in [('col',col),('colOff',0),('row',row),('rowOff',1)]:
                    ET.SubElement(edge, f'{{{X}}}{name}').text=str(val)
        attrs = {'Id':f'rId{i}', 'Type':f'{R}/oleObject', 'Target':f'../embeddings/report{i}.pdf'}
        if external: attrs.update(Target='https://example.invalid/report.pdf', TargetMode='External')
        ET.SubElement(rels, f'{{{P}}}Relationship', attrs)
        files[f'xl/embeddings/report{i}.pdf'] = pdf_bytes()
    files[key] = ET.tostring(root)
    files[f'xl/worksheets/_rels/sheet{sheet_no}.xml.rels'] = ET.tostring(rels)
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for n, data in files.items(): z.writestr(n,data)
    return path


def test_windows_embedded_name_decodes_on_mac():
    name = 'C:\\资料\\检测报告.pdf'.encode('gb18030')
    pdf = pdf_bytes()
    blob = struct.pack('<I',len(pdf)) + b'\x02\x00' + name + b'\x00' + b'ignored\x00' + pdf
    actual, data = _parse_ole10native(blob)
    assert actual == '检测报告.pdf'
    assert data == pdf


def test_encrypted_source_has_actionable_error_and_no_outputs(tmp_path, template):
    src = tmp_path/'RR10000.xlsx'; src.write_bytes(b'%TSD-Header-###%' + b'\0'*200)
    out = tmp_path/'output'
    res = Pipeline(str(src),str(template),str(out)).run()
    assert res.errors and '加密' in ' '.join(res.errors) and '解密' in ' '.join(res.errors)
    assert not list(out.glob('*.xlsx'))


@pytest.mark.parametrize('other_sheet', [False,True])
def test_true_xlsx_reads_selected_sheet_and_embedded_pdf(tmp_path, other_sheet):
    path = source_fixture(tmp_path/'RR10000.xlsx', other_sheet=other_sheet)
    sheet = SourceSheet(str(path),Settings())
    assert sheet.text(7,3) == 'Leadframe(model)'
    assert sheet.distinct_cells(10,7,8) == [(7,8)]
    objs = extract_embedded_objects(str(path))
    assert [(o.col,o.row,o.bottom_row) for o in objs] == [(10,7,8),(12,7,8)]
    assert all(o.data.startswith(b'%PDF') for o in objs)


@pytest.mark.parametrize('kind',['missing_anchor','external'])
def test_unassignable_objects_never_silently_drop(tmp_path,kind):
    p=source_fixture(tmp_path/'RR10000.xlsx',**{kind:True})
    with pytest.raises(ValueError): extract_embedded_objects(str(p))


def test_xlsx_content_with_xls_extension(tmp_path):
    p=source_fixture(tmp_path/'RR10000.xls')
    assert SourceSheet(str(p),Settings()).number(8,7) == 3


def test_rounding_total():
    comps=[Component(i,'a',str(i),1) for i in range(3)]
    assert normalize_content(comps) == [33.34,33.33,33.33]


def test_pipeline_using_template(tmp_path, template):
    p=source_fixture(tmp_path/'RR10000.xlsx')
    result=Pipeline(str(p),str(template),str(tmp_path/'out')).run()
    assert not result.errors
    assert (result.n_xlsx,result.n_hsf,result.n_msds)==(1,1,1)
    w=openpyxl.load_workbook(result.records[0].out_xlsx)
    s=w.worksheets[0]
    assert s['B2'].value == 'RR10000-Leadframe'
    assert [s['E6'].value,s['E7'].value] == [25,75]
    assert s['C18'].value == 'CANEC2612345678'
    assert s['B15'].value == 'ND'


def test_expansion_and_template_dropdown_preserved(tmp_path, template):
    rec=MaterialRecord(name='TEST-ONLY',components=[ComponentRow('Material',f'7440-{i}-1',12.5) for i in range(8)],
                       reports=[ReportRow('SGS',f'REP{i}','2026-10-01') for i in range(6)])
    out=tmp_path/'expanded.xlsx'
    build_workbook(str(template),str(out),rec,Settings())
    w=openpyxl.load_workbook(out); s=w.worksheets[0]
    assert sum(s.cell(r,5).value for r in range(6,14)) == 100
    assert s['A14'].value == '该行不允许删除'
    assert s['C25'].value == 'REP5'
    assert s['A26'].value == '该行不允许删除'
    assert 'A15:AC15' in str(s.merged_cells)
    with ZipFile(out) as z:
        assert b'dataValidations' in z.read('xl/worksheets/sheet1.xml')
        root=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        assert any(e.text=='H2' for e in root.iter() if e.tag.endswith('}sqref'))
    assert 'D6:D13' in str(s.data_validations)


def test_names_do_not_overwrite_existing_suffix_or_cleaned_name():
    names=_dedup(['Die','Die','Die-1','A/B','A:B','die'])
    from hsf.util import safe_name
    cleaned=[safe_name(n).casefold() for n in names]
    assert len(set(cleaned)) == len(names)


def test_wrong_template_rejected_before_output(tmp_path):
    src=source_fixture(tmp_path/'RR10000.xlsx')
    w=openpyxl.Workbook(); template=tmp_path/'bad.xlsx'; w.save(template)
    res=Pipeline(str(src),str(template),str(tmp_path/'out')).run()
    assert res.errors and not list((tmp_path/'out').glob('*'))


def test_failed_material_is_not_reported_as_success(tmp_path,monkeypatch,template):
    src=source_fixture(tmp_path/'RR10000.xlsx')
    def fail(*args,**kwargs): raise OSError('test: disk full')
    monkeypatch.setattr('hsf.pipeline.build_workbook',fail)
    res=Pipeline(str(src),str(template),str(tmp_path/'out')).run()
    assert res.errors and 'Leadframe' in ' '.join(res.errors)
    assert res.n_xlsx==0


def test_corrupt_pdf_not_exported_as_success(tmp_path):
    from hsf.pdfout import save_pdf_group
    from hsf.ole_extract import EmbeddedObject
    obj=EmbeddedObject(1,10,7,'broken.pdf',b'%PDF broken')
    path, warnings=save_pdf_group(str(tmp_path),'RR10000','HSF','Material',[obj])
    assert not path and warnings
    assert not list(tmp_path.glob('*.pdf'))


def test_explicit_chip_id_wins(tmp_path, template):
    p=Pipeline('RR10000-01.xlsx',str(template),str(tmp_path),prefix_override='RR10000-02')
    assert p._resolve_prefix()=='RR10000-02'


def test_placeholder_report_number_is_marked(tmp_path, template):
    rec=MaterialRecord(name='TEST',components=[ComponentRow('Cu','7440-50-8',100)],
                       reports=[ReportRow('SGS','TBD00000000','2026-10-01')])
    out=tmp_path/'placeholder.xlsx'
    build_workbook(str(template),str(out),rec,Settings())
    w=openpyxl.load_workbook(out)
    assert w.worksheets[0]['C18'].fill.fgColor.rgb.endswith('FFC7CE')


def ole_with_stream(name, data):
    # Minimal CFB v3 fixture with one regular (>4096-byte) stream.
    size=max(4096,len(data)); count=(size+511)//512
    assert count<120
    data=data.ljust(count*512,b'\0')
    FREE,END,FAT=0xFFFFFFFF,0xFFFFFFFE,0xFFFFFFFD
    header=bytearray(512); header[:8]=bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHHHH',header,24,0x3e,3,0xfffe,9,6)
    struct.pack_into('<IIIIIIIII',header,40,0,1,0,0,4096,END,0,END,0)
    struct.pack_into('<109I',header,76,count+1,*([FREE]*108))
    def entry(title,kind,child,start,length):
        e=bytearray(128); encoded=(title+'\0').encode('utf-16-le'); e[:len(encoded)]=encoded
        struct.pack_into('<HBBIII',e,64,len(encoded),kind,1,FREE,FREE,child)
        struct.pack_into('<IQ',e,116,start,length)
        return e
    directory=entry('Root Entry',5,1,END,0)+entry(name,2,FREE,1,size)+bytearray(256)
    sectors=[END]+[i+1 for i in range(1,count)]+[END,FAT]
    fat=struct.pack('<128I',*(sectors+[FREE]*(128-len(sectors))))
    return bytes(header)+directory+data+fat


@pytest.mark.parametrize('stream',['CONTENTS','\x01Ole10Native'])
def test_real_ole_wrapper_in_xlsx(tmp_path,stream):
    p=source_fixture(tmp_path/'RR10000.xlsx'); pdf=pdf_bytes()
    native=struct.pack('<I',len(pdf))+b'\x02\x00report.pdf\x00source\x00'+pdf
    wrapped=ole_with_stream(stream,pdf if stream=='CONTENTS' else native)
    with ZipFile(p) as z: entries={n:z.read(n) for n in z.namelist()}
    entries['xl/embeddings/report1.pdf']=wrapped
    with ZipFile(p,'w',ZIP_DEFLATED) as z:
        for n,data in entries.items(): z.writestr(n,data)
    from pypdf import PdfReader
    objects=extract_embedded_objects(str(p))
    assert 'SGS Test Report' in PdfReader(io.BytesIO(objects[0].data)).pages[0].extract_text()


def test_legacy_xls_table_and_disguised_extension(tmp_path):
    import xlwt
    p=tmp_path/'RR10000.xlsx'
    w=xlwt.Workbook(); s=w.add_sheet('材料成分表')
    s.write(7,3,'Copper'); s.write(7,5,'Copper'); s.write(7,6,'7440-50-8'); s.write(7,7,1)
    s.write_merge(7,8,10,10,'HSF'); w.save(str(p))
    sheet=SourceSheet(str(p),Settings())
    assert sheet.text(7,3)=='Copper' and sheet.distinct_cells(10,7,8)==[(7,8)]
    assert extract_embedded_objects(str(p))==[]
