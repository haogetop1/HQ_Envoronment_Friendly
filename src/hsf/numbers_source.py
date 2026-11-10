"""Read Numbers tables and resolve their floating attachment previews locally."""
from bisect import bisect_right
from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import datetime, time
from io import BytesIO
from itertools import accumulate
from pathlib import Path
import re
from shutil import copyfile, copytree
from subprocess import run
from tempfile import TemporaryDirectory
import unicodedata

from numbers_parser import Document
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Font
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
                previous = []
                for row in range(min(20, table.num_rows)):
                    values = _headers(table, merges, row)
                    columns = columns_for_headers(values, settings.columns)
                    if columns is None and previous:
                        # Some imports leave the report label one row above the Chinese header.
                        with_report = [
                            '检测报告' if not value and re.fullmatch(
                                r'(?:sgs|hsf|rohs)(?:检测报告)?|检测报告',
                                re.sub(r'\s+', '', str(above or '')).casefold()) else value
                            for value, above in zip(values, previous)
                        ]
                        columns = columns_for_headers(with_report, settings.columns)
                    if columns is not None:
                        candidates.append((sheet, table, merges, row, columns))
                        break
                    previous = values
        preferred = [entry for entry in candidates if entry[0].name == settings.sheet_name]
        candidates = preferred or candidates
        if len(candidates) != 1:
            raise ValueError('Numbers 中未找到唯一的材料成分表。请保留一张含材料型号、成分、CAS、重量、检测报告和 MSDS 表头的表。')
        sheet, self.table, self.merged_cells, row, columns = candidates[0]
        self.name = sheet.name
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
                        yield obj_id, left, top, geo.size.width, geo.size.height, name, data

        yield from walk(self.drawables)

    def attachments(self, source, pdf_dir=None):
        rows, cols = self._grid()
        columns = self.settings.columns
        files = _pdf_index(pdf_dir)
        start, end = source.detect_data_range()
        attachments = []
        for obj_id, left, top, width, height, name, data in self._pdf_drawings():
            # Floating Numbers images have no Excel anchors. Use the center of the drawing.
            col = bisect_right(cols, left + width / 2) - 1
            row = bisect_right(rows, top + height / 2) - 1
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


def _add_pdf_previews(sheet, numbers_sheet, directory):
    """Embed visible drawings as PNGs, as Numbers does when exporting to Excel."""
    drawings = list(numbers_sheet._pdf_drawings())
    if not drawings:
        return
    rows, cols = numbers_sheet._grid()
    for obj_id, x, y, width, height, name, data in drawings:
        if not data or not data.startswith(b'%PDF-') or width <= 0 or height <= 0:
            raise ValueError(f'Numbers PDF 对象无法导出：{name}')
        pdf = directory / f'{obj_id}.pdf'
        png = pdf.with_suffix('.png')
        pdf.write_bytes(data)
        result = run(['/usr/bin/sips', '-s', 'format', 'png', str(pdf), '--out', str(png)],
                     capture_output=True, timeout=30)
        if result.returncode or not png.is_file():
            raise ValueError(f'Numbers PDF 对象无法渲染：{name}')
        x, y = max(x, cols[0]), max(y, rows[0])
        col = min(bisect_right(cols, x) - 1, len(cols) - 2)
        row = min(bisect_right(rows, y) - 1, len(rows) - 2)
        # Numbers geometry is in points; Excel drawing coordinates are EMUs.
        image = Image(BytesIO(png.read_bytes()))
        image.anchor = OneCellAnchor(
            _from=AnchorMarker(col=col, row=row, colOff=round((x - cols[col]) * 12700),
                               rowOff=round((y - rows[row]) * 12700)),
            ext=XDRPositiveSize2D(round(width * 12700), round(height * 12700)))
        sheet.add_image(image)


def write_converted_xlsx(path, numbers_sheet):
    """Copy the material table and its PDF object previews at their original coordinates."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = re.sub(r'[\\/*?:\[\]]', '_', numbers_sheet.name)[:31] or '材料成分表'
    table = numbers_sheet.table
    date_columns = {
        col for row in range(numbers_sheet.settings.header_row + 1)
        for col, value in enumerate(_headers(table, numbers_sheet.merged_cells, row))
        if re.search(r'testdate|检测时间|检测日期', re.sub(r'\s+', '', str(value or '')).casefold())
    }
    date_widths = {}
    wrapped = Alignment(wrap_text=True, vertical='center')
    header = Font(bold=True)
    for row in range(numbers_sheet.nrows):
        sheet.row_dimensions[row + 1].height = table.row_height(row)
        for col in range(numbers_sheet.ncols):
            native = table.cell(row, col)
            value = native.value
            if value is None:
                continue
            cell = sheet.cell(row + 1, col + 1, value)
            if isinstance(value, str):
                cell.data_type = 's'
                cell.alignment = wrapped
            elif col in date_columns and isinstance(value, datetime):
                has_time = value.time() != time.min
                cell.number_format = 'yyyy/mm/dd hh:mm:ss' if has_time else 'yyyy/mm/dd'
                date_widths[col] = max(date_widths.get(col, 0), 25 if has_time else 16)
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                if str(native.formatted_value).rstrip().endswith(('%', '％')):
                    cell.number_format = '0.00%'
            if row <= numbers_sheet.settings.header_row:
                cell.font = header
    for col in range(numbers_sheet.ncols):
        sheet.column_dimensions[get_column_letter(col + 1)].width = max(
            table.col_width(col) / 5.25, date_widths.get(col, 0))
    for top, bottom, left, right in numbers_sheet.merged_cells:
        sheet.merge_cells(start_row=top + 1, end_row=bottom, start_column=left + 1, end_column=right)
    sheet.freeze_panes = sheet.cell(numbers_sheet.settings.data_start_row + 1, 1)
    path = Path(path)
    with TemporaryDirectory(prefix='hsf-export-', dir=path.parent) as temp:
        directory = Path(temp)
        _add_pdf_previews(sheet, numbers_sheet, directory)
        staged = directory / 'converted.xlsx'
        workbook.save(staged)
        staged.replace(path)


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
