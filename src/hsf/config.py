# -*- coding: utf-8 -*-
"""固定列位配置与业务常量。

列位一律使用 0-based 索引（A=0, B=1 ... K=10, M=12）。
本工具面向「相同架构」的 .xls 文件：列位固定，但全部可在界面/调用处覆盖。
"""
from dataclasses import dataclass, field


@dataclass
class Columns:
    """源 .xls（材料成分表）的列位。"""
    material_name: int = 2      # C 列：材料名称（本项目按用户确认为「材料厂家」列）
    material_model: int = 3     # D 列：材料名称/型号
    component: int = 5          # F 列：材料成分（物质名称）
    cas: int = 6                # G 列：CAS No.
    mass: int = 7               # H 列：材料成份重量（归一化基准）
    hsf: int = 10               # K 列：HSF / SGS 检测报告（内嵌 PDF）
    msds: int = 12              # M 列：MSDS（内嵌 PDF）


@dataclass
class Settings:
    sheet_name: str = '材料成分表'
    header_row: int = 6         # 0-based，中文表头所在行（Excel 第 7 行）
    data_start_row: int = 7     # 0-based，数据首行
    columns: Columns = field(default_factory=Columns)

    # 材料名称取自哪一列：'model'=D列(材料名称/型号) 'vendor'=C列(材料厂家) 'both'=C+D
    name_source: str = 'model'

    # CAS 为空时是否判定为保密
    secret_when_cas_empty: bool = True

    # 固定填写值
    supplier: str = 'SAC'
    model: str = 'TBD'
    eco_level: str = 'Ⅰ级(无卤欧盟环保)'
    category: str = 'Metal'
    color: str = '黑色'
    # 2026-08-31：已废弃。抓不到编号时改为弹窗询问，不再用固定兜底值，
    # 避免换了芯片还静默套用旧编号。保留字段仅为兼容旧配置文件。
    fallback_prefix: str = ''

    # 数值
    content_decimals: int = 2

    # 模板固定结构（1-based 行号）
    tpl_row_name: int = 2
    tpl_row_cat: int = 3
    tpl_comp_header: int = 5
    tpl_comp_first: int = 6
    tpl_comp_marker: int = 12      # 「该行不允许删除」守护行
    tpl_result_header: int = 14
    tpl_result_row: int = 15
    tpl_report_header: int = 17
    tpl_report_first: int = 18
    tpl_report_marker: int = 22    # 「该行不允许删除」守护行

    comp_capacity: int = 6         # 6..11
    report_capacity: int = 4       # 18..21


# 模板第 14 行的测试项目顺序 -> 列号（1-based）
RESULT_ITEMS = [
    ('Pb', 2), ('Cd', 3), ('Hg', 4), ('Cr6+', 5),
    ('PBBs', 6), ('PBDEs', 7), ('BBP', 8), ('DBP', 9), ('DEHP', 10), ('DIBP', 11),
    ('Br', 12), ('Cl', 13), ('F', 14),
    ('HBCDD', 15), ('SCCP', 16), ('PFOA', 17), ('PFOS', 18),
    ('MOSH/MOAH', 19), ('VOC', 20), ('Formaldehyde', 21),
    ('Sb', 22), ('Be', 23), ('As', 24), ('P', 25),
    ('DINP', 26), ('DnHP', 27), ('DIDP', 28), ('DNOP', 29),
]

# 必填项：不允许 NA
MANDATORY_ITEMS = ['Pb', 'Cd', 'Hg', 'Cr6+']

# 领域规则：同一组的项目必然一起测试。
# 组内只要有一项测出结果，其余仍为 NA 的项即判定为「抽取遗漏」并补 ND。
DOMAIN_GROUPS = [
    ['F', 'Cl', 'Br'],
    ['PBBs', 'PBDEs', 'BBP', 'DBP', 'DEHP', 'DIBP'],
    ['DINP', 'DnHP', 'DIDP', 'DNOP'],
    ['PFOA', 'PFOS'],
]

# 测试机构关键字 -> 输出简称（按顺序匹配，先命中者胜）
# 2026-08-31 补 HCT（深圳市沃特虹彩检测 / Waltek Testing Group HCT）。
# 关键字取自报告页脚的 HCT/RF-23-19 表单号与 hct-test.com 域名——比裸 'HCT'
# 更不容易误伤其它报告正文里恰好含 hct 的英文单词。
AGENCY_RULES = [
    (['Centre Testing International', '華測', '华测', 'CTI '], '华测'),
    (['Meixin', '美信', 'mttlab'], '美信'),
    (['hct-test', 'Waltek', '虹彩', 'HCT/RF'], 'HCT'),
    (['SGS'], 'SGS'),
    (['TÜV', 'TUV'], 'TÜV'),
    (['Intertek', '天祥'], 'Intertek'),
    (['BV ', 'Bureau Veritas'], '必维'),
]

# 2026-08-31 需求：HSF 报告编号识别不出时的占位值。
FALLBACK_REPORT_NO = 'TBD00000000'

# 物质名称 -> PDF 中可能出现的标签（正则，忽略大小写）
ITEM_LABELS = {
    'Pb': [r'Lead\s*\(Pb\)', r'\bLead\b', r'\bPb\b'],
    'Cd': [r'Cadmium\s*\(Cd\)', r'Cadmium\(Cd\)', r'\bCadmium\b'],
    'Hg': [r'Mercury\s*\(Hg\)', r'\bMercury\b', r'\bHg\b'],
    'Cr6+': [r'Hexavalent\s*Chromium\s*\(?Cr\(VI\)|Cr6\+?\)?', r'Hexavalent\s+Chromium',
             r'Hexavalent\s*Chromium\(Cr6\+\)', r'Cr\(VI\)', r'Cr6\+'],
    'PBBs': [r'Sum\s+of\s+PBBs?', r'Polybrominated\s+biphenyls?\s*\(PBB\)', r'PBBs?'],
    'PBDEs': [r'Sum\s+of\s+PBDEs?', r'Polybrominated\s+diphenyl\s+ethers?\s*\(PBDE\)', r'PBDEs?'],
    'BBP': [r'Butyl\s+benzyl\s+phthalate\s*\(BBP\)', r'\bBBP\b'],
    'DBP': [r'Dibutyl\s+[Pp]hthalate\s*\(DBP\)', r'\bDBP\b'],
    'DEHP': [r'Bis\s*\(2-ethylhexyl\)\s*phthalate\s*\(DEHP\)', r'Bis\(2-ethylhexyl\)\s*phthalate', r'\bDEHP\b'],
    'DIBP': [r'Diisobutyl\s+[Pp]hthalate\s*\(DIBP\)', r'\bDIBP\b'],
    'Br': [r'Bromine\s*\(Br\)', r'\bBromine\b'],
    'Cl': [r'Chlorine\s*\(Cl\)', r'\bChlorine\b'],
    'F': [r'Fluorine\s*\(F\)', r'\bFluorine\b'],
    'HBCDD': [r'Hexabromocyclododecane', r'\bHBCDD\b'],
    'SCCP': [r'Short\s+[Cc]hain\s+[Cc]hlorinated\s+[Pp]araffins', r'\bSCCP\b', r'Alkanes,?\s*C10-13'],
    'PFOA': [r'Perfluorooctanoic\s+acid\s*\(PFOA\)', r'\bPFOA\b'],
    'PFOS': [r'Perfluorooctane\s+sulfonic?\s+acid\s*\(PFOS\)', r'\bPFOS\b'],
    'MOSH/MOAH': [r'MOSH', r'MOAH', r'Mineral\s+[Oo]il'],
    'VOC': [r'\bVOC\b', r'Volatile\s+[Oo]rganic\s+[Cc]ompound'],
    'Formaldehyde': [r'\bFormaldehyde\b'],
    'Sb': [r'Antimony\s*\(Sb\)', r'\bAntimony\b'],
    'Be': [r'Beryllium\s*\(Be\)', r'\bBeryllium\b'],
    'As': [r'Arsenic\s*\(As\)', r'\bArsenic\b'],
    'P': [r'Red\s+[Pp]hosphorus', r'^\s*Phosphorus\s*\(P\)'],
    'DINP': [r'Diisononyl\s+[Pp]hthalate\s*\(DINP\)', r'\bDINP\b'],
    'DnHP': [r'Di-n-[Hh]exyl\s+[Pp]hthalate\s*\(DnHP\)', r'\bDnHP\b', r'Di-n-hexyl\s+phthalate'],
    'DIDP': [r'Diisodecyl\s+[Pp]hthalate\s*\(DIDP\)', r'\bDIDP\b'],
    'DNOP': [r'Di-n-[Oo]ctyl\s+[Pp]hthalate\s*\(DNOP\)', r'\bDNOP\b'],
}

# 判定「保密」的 CAS 关键字
SECRET_CAS = {'trade secret', 'proprietary', 'secret', 'confidential', '商业机密', '机密'}

# 2026-08-31 需求：无 CAS NO 的保密物质，CAS 栏统一填这个占位号
SECRET_CAS_VALUE = '00-00-00'
