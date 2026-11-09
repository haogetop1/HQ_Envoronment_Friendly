"""Read OOXML embedded objects and anchors without executing Office objects."""
import io
import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile
import olefile

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'


def relationships(z, part):
    parent, name = posixpath.split(part)
    relfile = posixpath.join(parent, '_rels', name + '.rels')
    if relfile not in z.namelist():
        return {}
    return {e.get('Id'):e.attrib for e in ET.fromstring(z.read(relfile))}


def resolve(part, rel):
    if rel.get('TargetMode') == 'External':
        raise ValueError('报告仅为外部链接，未内嵌到 Excel；请重新插入实际 PDF 后重试。')
    target = rel.get('Target','')
    path = posixpath.normpath(posixpath.join(posixpath.dirname(part),target)) if not target.startswith('/') else target[1:]
    if path.startswith('../') or ':' in path or not path:
        raise ValueError('Excel 包含无效的内嵌对象路径。')
    return path


def sheet_part(z, name):
    wb = 'xl/workbook.xml'
    sheets = ET.fromstring(z.read(wb)).find(f'{{{S}}}sheets')
    if sheets is None or not len(sheets):
        raise ValueError('Excel 中没有工作表。')
    chosen = next((e for e in sheets if e.get('name')==name),sheets[0])
    return resolve(wb,relationships(z,wb)[chosen.get(f'{{{R}}}id')])


def anchor_coords(anchor):
    edges = {e.tag.split('}')[-1]:e for e in anchor}
    def value(edge, tag):
        return next(int(e.text) for e in edges[edge] if e.tag.split('}')[-1]==tag)
    try:
        col, top, bottom = value('from','col'),value('from','row'),value('to','row')
    except (KeyError,StopIteration,ValueError,TypeError):
        raise ValueError('无法读取报告图标的单元格位置，请将报告重新嵌入到对应列。')
    if min(col,top,bottom)<0 or bottom<top or col>=16384 or bottom>=1048576:
        raise ValueError('报告图标位置超出 Excel 范围。')
    return col,top,bottom


def drawing_anchors(z, part, rels):
    anchors = {}
    for rel in rels.values():
        if rel.get('Type','').endswith('/drawing'):
            root = ET.fromstring(z.read(resolve(part,rel)))
            for anchor in root:
                if anchor.tag.split('}')[-1] != 'twoCellAnchor':
                    continue
                for node in anchor.iter():
                    if node.tag.split('}')[-1]=='cNvPr' and node.get('id'):
                        anchors[node.get('id')] = anchor_coords(anchor)
        elif rel.get('Type','').endswith('/vmlDrawing'):
            root = ET.fromstring(z.read(resolve(part,rel)))
            for shape in root.iter():
                if shape.tag.split('}')[-1]!='shape':
                    continue
                oid = (shape.get('id') or '').rsplit('s',1)[-1]
                for node in shape.iter():
                    if node.tag.split('}')[-1]=='Anchor' and node.text:
                        values = [int(v.strip()) for v in node.text.split(',')]
                        if len(values)==8:
                            anchors[oid] = (values[0],values[2],values[6])
    return anchors


def extract_xlsx(path, sheet_name='材料成分表'):
    from .ole_extract import EmbeddedObject, _parse_ole10native
    result = []
    with ZipFile(path) as z:
        part = sheet_part(z,sheet_name)
        root = ET.fromstring(z.read(part))
        rels = relationships(z,part)
        fallback = drawing_anchors(z,part,rels)
        seen = set()
        for obj in root.iter(f'{{{S}}}oleObject'):
            rid = obj.get(f'{{{R}}}id')
            if rid in seen:
                continue
            seen.add(rid)
            if rid not in rels:
                raise ValueError('内嵌对象缺少文件关系，无法安全提取。')
            target = resolve(part,rels[rid])
            anchor = obj.find(f'{{{S}}}objectPr/{{{S}}}anchor')
            coords = anchor_coords(anchor) if anchor is not None else fallback.get(obj.get('shapeId'))
            if coords is None:
                raise ValueError(f'内嵌报告 {target} 没有可识别的单元格锚点，不能猜测材料归属。')
            data = z.read(target); filename = posixpath.basename(target)
            if olefile.isOleFile(io.BytesIO(data)):
                with olefile.OleFileIO(io.BytesIO(data)) as ole:
                    streams = ole.listdir()
                    native = next((p for p in streams if p[-1].endswith('Ole10Native')),None)
                    contents = next((p for p in streams if p[-1].upper()=='CONTENTS'),None)
                    if native:
                        filename,data = _parse_ole10native(ole.openstream(native).read())
                    elif contents:
                        data = ole.openstream(contents).read()
                        filename = posixpath.basename(target).rsplit('.',1)[0]+'.pdf'
                    else:
                        raise ValueError(f'无法提取内嵌对象 {target}，请将其另存为 PDF 后重新嵌入。')
            if not data.startswith(b'%PDF'):
                raise ValueError(f'内嵌对象 {filename} 不是 PDF，请转换为 PDF 后重新嵌入。')
            col,top,bottom = coords
            result.append(EmbeddedObject(int(obj.get('shapeId','0')),col,top,filename,data,bottom))
    return result
