"""Read Numbers tables and resolve their floating attachment previews locally."""
from bisect import bisect_right
from collections import defaultdict
from dataclasses import dataclass, replace
from io import BytesIO
from itertools import accumulate
from pathlib import Path
import re
from shutil import copyfile, copytree
from tempfile import TemporaryDirectory
import unicodedata

from numbers_parser import Document
from openpyxl import Workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from pypdf import PdfReader

from .ole_extract import EmbeddedObject
from .xls_source import columns_for_headers


def _open_document(path):
    path = Path(path)
    if path.suffix == '.numbers':
        return Document(path)
    with TemporaryDirectory(prefix='hsf-numbers-') as temp:
        copy = Path(temp) / 'source.numbers'
        (copytree if path.is_dir() else copyfile)(path, copy)
        return Document(copy)


def _merge_spans(table):
    for ref in table.merge_ranges:
        first_col, first_row, last_col, last_row = range_boundaries(ref)
        yield first_row - 1, last_row, first_col - 1, last_col


def _headers(table, merges, row):
    values = [table.cell(row, col).value for col in range(table.num_cols)]
    for top, bottom, left, right in merges:
        if top <= row < bottom:
            for col in range(left, right):
                values[col] = table.cell(top, left).value
    return values


class NumbersSheet:
    def __init__(self, path, settings):
        document = _open_document(path)
        candidates = []
        for sheet in document.sheets:
            for table in sheet.tables:
                merges = list(_merge_spans(table))
                for row in range(min(20, table.num_rows)):
                    columns = columns_for_headers(_headers(table, merges, row), settings.columns)
                    if columns is not None:
                        candidates.append((sheet, table, merges, row, columns))
                        break
        preferred = [entry for entry in candidates if entry[0].name == settings.sheet_name]
        candidates = preferred or candidates
        if len(candidates) != 1:
            raise ValueError('Numbers 中未找到唯一的材料成分表。请保留一张含材料型号、成分、CAS、重量、检测报告和 MSDS 表头的表。')
        sheet, self.table, self.merged_cells, row, columns = candidates[0]
        self.settings = replace(settings, header_row=row, data_start_row=row + 1, columns=columns)
        self.nrows, self.ncols = self.table.num_rows, self.table.num_cols
        # Pinned parser API: the public table interface does not expose drawable geometry.
        self.store = document._model.objects
        self.drawables = self.store[sheet._sheet_id].drawable_infos

    def cell_value(self, row, col):
        value = self.table.cell(row, col).value
        return '' if value is None else value

    def _grid(self):
        model = self.store[self.table._table_id]
        info = next(self.store[ref.identifier] for ref in self.drawables
                    if ref.identifier in self.store
                    and self.store[ref.identifier].DESCRIPTOR.name == 'TableInfoArchive'
                    and self.store[ref.identifier].tableModel.identifier == self.table._table_id)
        geometry = info.super.geometry
        if geometry.angle:
            raise ValueError('Numbers 材料表已旋转，无法定位报告。')
        bds = model.base_data_store
        heights = {entry.index: entry.size for ref in bds.rowHeaders.buckets
                   for entry in self.store[ref.identifier].headers}
        widths = {entry.index: entry.size for entry in self.store[bds.columnHeaders.identifier].headers}
        fitted = info.layout_engine.width_height_cache.rows_fitting_entries

        def automatic_height(row):
            size = fitted[row].fitting_size if row < len(fitted) else 0
            return size or model.default_row_height

        rows = list(accumulate((heights.get(r) or automatic_height(r) for r in range(self.nrows)),
                               initial=geometry.position.y))
        cols = list(accumulate((widths.get(c) or model.default_column_width for c in range(self.ncols)),
                               initial=geometry.position.x))
        return rows, cols

    def _pdf_drawings(self):
        package = self.store[self.store.find_refs('PackageMetadata')[0]]
        filenames = {ref.identifier: ref.file_name for ref in package.datas}
        seen = set()

        def walk(references, x=0, y=0):
            for ref in references:
                obj_id = ref.identifier
                if obj_id in seen or obj_id not in self.store:
                    continue
                seen.add(obj_id)
                obj = self.store[obj_id]
                kind = obj.DESCRIPTOR.name
                if kind not in ('GroupArchive', 'ImageArchive'):
                    continue
                geo = obj.super.geometry
                left, top = x + geo.position.x, y + geo.position.y
                if kind == 'GroupArchive':
                    yield from walk(obj.children, left, top)
                else:
                    name = filenames.get(obj.data.identifier, '')
                    if name.lower().endswith('.pdf'):
                        data = self.store.file_store.get('Data/' + name)
                        yield obj_id, left + geo.size.width / 2, top + geo.size.height / 2, name, data

        yield from walk(self.drawables)

    def attachments(self, source, pdf_dir=None):
        rows, cols = self._grid()
        columns = self.settings.columns
        files = _pdf_index(pdf_dir)
        start, end = source.detect_data_range()
        attachments = []
        for obj_id, x, y, name, data in self._pdf_drawings():
            # Floating Numbers images have no Excel anchors. Use the center of the drawing.
            col, row = bisect_right(cols, x) - 1, bisect_right(rows, y) - 1
            if col not in (columns.hsf, columns.msds) or not start <= row <= end:
                continue
            kind = 'HSF' if col == columns.hsf else 'MSDS'
            top, bottom = next((a, b) for a, b in source.distinct_cells(col, start, end)
                               if a <= row <= b)
            label = _preview_label(data)
            pdf, status = _resolve_pdf(files, label)
            attachments.append(Attachment(kind, col, top, bottom, label or name, status,
                                          EmbeddedObject(obj_id, col, row, label, pdf) if pdf else None))
        return attachments


@dataclass
class Attachment:
    kind: str
    col: int
    start: int
    end: int
    label: str
    status: str
    object: EmbeddedObject | None


def _preview_label(data):
    if not data:
        return ''
    try:
        pdf = PdfReader(BytesIO(data))
        if len(pdf.pages) != 1:
            return ''
        text = re.sub(r'\s*[\r\n]+\s*', '', pdf.pages[0].extract_text()).strip()
    except Exception:
        return ''
    text = unicodedata.normalize('NFKC', text)
    if len(text) <= 240 and re.fullmatch(r'.+\.(pdf|xlsx?|docx?)', text, re.I):
        return text
    return ''


def _filename_key(name):
    # Names rendered into the icon can wrap or use different Unicode whitespace.
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', name)).casefold()


def _pdf_index(directory):
    indexed = defaultdict(list)
    if directory:
        path = Path(directory)
        if not path.is_dir():
            raise ValueError('指定的原始报告文件夹不存在。')
        for file in path.rglob('*'):
            if file.suffix.lower() == '.pdf' and file.is_file() and not file.is_symlink():
                indexed[_filename_key(file.name)].append(file)
    return indexed


def _resolve_pdf(files, label):
    if not label:
        return None, '图标中无法识别原始文件名'
    if not label.lower().endswith('.pdf'):
        return None, '原附件不是 PDF'
    matches = files.get(_filename_key(label), [])
    if not matches:
        return None, '缺少原始 PDF'
    if len(matches) > 1:
        return None, '有多个同名 PDF，无法确定'
    try:
        data = matches[0].read_bytes()
        reader = PdfReader(BytesIO(data))
        if not data.startswith(b'%PDF-') or not len(reader.pages) or _preview_label(data):
            raise ValueError('不是原始报告')
    except Exception:
        return None, '同名 PDF 无法读取或仍为图标预览'
    return data, '已匹配原始 PDF'


def write_audit(path, source, attachments):
    workbook = Workbook()
    review = workbook.active
    review.title = '附件核对'
    review.append(['类型', '源表区域', '材料名称', '图标文件名', '状态'])
    c = source.st.columns
    for item in attachments:
        first, last = item.start + 1, item.end + 1
        letter = get_column_letter(item.col + 1)
        area = f'{letter}{first}:{letter}{last}' if first != last else f'{letter}{first}'
        review.append([item.kind, area, source.text(item.start, c.material_model), item.label, item.status])
    materials = workbook.create_sheet('材料成分')
    materials.append(['源表行号', '材料名称', '物质名称', 'CAS No.', '原表重量'])
    start, end = source.detect_data_range()
    for row in range(start, end + 1):
        substance, cas = source.text(row, c.component), source.text(row, c.cas)
        if substance or cas:
            materials.append([row + 1, source.text(row, c.material_model), substance, cas,
                              source.cell(row, c.mass)])
    for sheet, widths in ((review, (12, 17, 26, 68, 34)), (materials, (13, 28, 36, 23, 20))):
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        for column, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(column)].width = width
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
    workbook.save(path)
