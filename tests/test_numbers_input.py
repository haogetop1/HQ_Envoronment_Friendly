import io
import hashlib
from pathlib import Path
import shutil

from numbers_parser import Document
from numbers_parser.generated.TSDArchives_pb2 import ImageArchive
from openpyxl import load_workbook
from reportlab.pdfgen.canvas import Canvas

from hsf.formats import workbook_format
from hsf.pipeline import Pipeline
from hsf.xls_source import SourceSheet
from hsf.config import Settings
from test_mac_port import pdf_bytes, template


def numbers_fixture(path: Path):
    doc = Document(sheet_name='材料成分表', num_rows=8, num_cols=13)
    table = doc.default_table
    for col, value in {3:'材料名称/型号', 5:'材料成分', 6:'CAS 编号', 7:'材料成份重量',
                       10:'SGS检测报告', 12:'MSDS'}.items():
        table.write(3, col, value)
    table.write(4, 3, 'Leadframe')
    for row, material, cas, mass in [(4, 'Copper', '7440-50-8', 1), (5, 'Tin', '7440-31-5', 3)]:
        for col, val in {5:material, 6:cas, 7:mass}.items():
            table.write(row, col, val)
    table.write(6, 0, '合计')
    table.merge_cells(['K5:K6','M5:M6'])
    for row in range(table.num_rows):
        table.row_height(row, 20)
    for col in range(table.num_cols):
        table.col_width(col, 100)
    doc.save(path)
    doc = Document(path)
    store = doc._model.objects
    sheet = store[doc.sheets[0]._sheet_id]
    info = store[store.find_refs('TableInfoArchive')[0]]
    position = info.super.geometry.position
    package = store[store.find_refs('PackageMetadata')[0]]
    for col, label in [(10, 'Report One.pdf'), (12, 'Safety Sheet.pdf')]:
        buf = io.BytesIO()
        canvas = Canvas(buf, pagesize=(120, 100))
        canvas.drawString(6, 60, label)
        canvas.save()
        data_id = 1000 + col
        filename = f'image-{data_id}.pdf'
        package.datas.add(identifier=data_id, file_name=filename, preferred_file_name=filename,
                          digest=hashlib.sha1(buf.getvalue()).digest())
        store.file_store['Data/'+filename] = buf.getvalue()
        obj_id, _ = store.create_object_from_dict('Index/Document', {
            'super': {'geometry': {'position': {'x': position.x+col*100+10, 'y': position.y+85},
                                  'size': {'width':40, 'height':20}}},
            'data': {'identifier':data_id},
        }, ImageArchive, append=True)
        sheet.drawable_infos.add(identifier=obj_id)
    store.update_object_file_store()
    store.save(path, False, None, None)


def test_numbers_reads_merges_and_header_with_alias(tmp_path):
    path = tmp_path/'RR10000.numbers'
    numbers_fixture(path)
    alias = tmp_path/'RR10000.number'
    shutil.copyfile(path, alias)
    assert workbook_format(str(alias)) == 'numbers'
    sheet = SourceSheet(str(alias), Settings())
    assert sheet.st.header_row == 3
    assert sheet.distinct_cells(10, 4, 5) == [(4, 5)]
    assert sheet.number(5, 7) == 3


def test_numbers_icon_does_not_become_report(tmp_path, template):
    source = tmp_path/'RR10000.numbers'
    numbers_fixture(source)
    output = tmp_path/'output'
    result = Pipeline(str(source), str(template), str(output)).run()
    assert result.errors and result.n_xlsx == 0
    audit = load_workbook(output/'Numbers附件核对.xlsx', data_only=True)
    assert list(audit['附件核对'].values)[1][3:5] == ('Report One.pdf', '缺少原始 PDF')


def test_numbers_sidecar_originals_fill_template(tmp_path, template):
    source = tmp_path/'RR10000.numbers'
    numbers_fixture(source)
    pdf_dir = tmp_path/'reports'
    pdf_dir.mkdir()
    (pdf_dir/'Report One.pdf').write_bytes(pdf_bytes(pb='17'))
    (pdf_dir/'Safety Sheet.pdf').write_bytes(pdf_bytes())
    output = tmp_path/'output'
    result = Pipeline(str(source), str(template), str(output), pdf_dir=str(pdf_dir)).run()
    assert not result.errors
    assert (result.n_xlsx, result.n_hsf, result.n_msds) == (1, 1, 1)
    sheet = load_workbook(result.records[0].out_xlsx, data_only=True).worksheets[0]
    assert [sheet['E6'].value, sheet['E7'].value] == [25, 75]
    assert sheet['C18'].value == 'CANEC2612345678'
    assert str(sheet['B15'].value) == '17'


def test_numbers_refuses_to_choose_between_same_named_reports(tmp_path, template):
    source = tmp_path/'RR10000.numbers'
    numbers_fixture(source)
    for folder in ('first', 'second'):
        directory = tmp_path/'reports'/folder
        directory.mkdir(parents=True)
        (directory/'Report One.pdf').write_bytes(pdf_bytes())
    output = tmp_path/'output'
    result = Pipeline(str(source), str(template), str(output), pdf_dir=str(tmp_path/'reports')).run()
    assert result.errors and result.n_xlsx == 0
    audit = load_workbook(output/'Numbers附件核对.xlsx', data_only=True)
    assert audit['附件核对']['E2'].value == '有多个同名 PDF，无法确定'
