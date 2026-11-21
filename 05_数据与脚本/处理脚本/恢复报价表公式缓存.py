#!/usr/bin/env python3
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx'
TARGET = ROOT / '来福商品池_成本及二次定价_回填上海履约成本_2026-11-02.xlsx'
TEMP = ROOT / '来福商品池_成本及二次定价_回填上海履约成本_缓存恢复.tmp.xlsx'
CELL_RE = re.compile(rb'<c\b[^>]*>.*?</c>', re.S)
REF_RE = re.compile(rb'\br="([A-Z]+[0-9]+)"')
TYPE_RE = re.compile(rb'\s+t="[^"]*"')
VALUE_RE = re.compile(rb'<v(?:\s[^>]*)?>.*?</v>|<v\s*/>', re.S)
FORMULA_RE = re.compile(rb'<f(?:\s|>)')


def formula_caches(xml):
    caches = {}
    for match in CELL_RE.finditer(xml):
        cell = match.group(0)
        if not FORMULA_RE.search(cell):
            continue
        opening_end = cell.find(b'>') + 1
        opening = cell[:opening_end]
        ref_match = REF_RE.search(opening)
        value_match = VALUE_RE.search(cell)
        assert ref_match and value_match, '公式单元格缺少引用或缓存节点'
        type_match = TYPE_RE.search(opening)
        caches[ref_match.group(1)] = (type_match.group(0) if type_match else b'', value_match.group(0))
    return caches


def restore(xml, caches, sheet_name):
    seen = set()
    def replace(match):
        cell = match.group(0)
        if not FORMULA_RE.search(cell):
            return cell
        opening_end = cell.find(b'>') + 1
        opening = cell[:opening_end]
        ref_match = REF_RE.search(opening)
        assert ref_match
        ref = ref_match.group(1)
        assert ref in caches, f'{sheet_name}!{ref.decode()} 缺少原缓存'
        assert ref not in seen, f'{sheet_name}!{ref.decode()} 重复公式单元格'
        seen.add(ref)
        cached_type, cached_value = caches[ref]
        new_opening = TYPE_RE.sub(b'', opening)
        new_opening = new_opening[:-1] + cached_type + b'>'
        new_cell = new_opening + cell[opening_end:]
        value_match = VALUE_RE.search(new_cell)
        assert value_match, f'{sheet_name}!{ref.decode()} 新文件缺少缓存节点'
        new_cell = new_cell[:value_match.start()] + cached_value + new_cell[value_match.end():]
        return new_cell
    restored = CELL_RE.sub(replace, xml)
    assert seen == set(caches), f'{sheet_name} 公式数量不一致：原{len(caches)}，新{len(seen)}'
    return restored, len(seen)


with zipfile.ZipFile(SOURCE) as source_zip, zipfile.ZipFile(TARGET) as target_zip, zipfile.ZipFile(TEMP, 'w') as temp_zip:
    source_names = set(source_zip.namelist())
    total = 0
    for info in target_zip.infolist():
        data = target_zip.read(info.filename)
        if info.filename.startswith('xl/worksheets/sheet') and info.filename.endswith('.xml'):
            assert info.filename in source_names
            source_xml = source_zip.read(info.filename)
            caches = formula_caches(source_xml)
            if caches or FORMULA_RE.search(data):
                data, count = restore(data, caches, info.filename)
                total += count
                print(f'{info.filename} 恢复公式缓存 {count}', flush=True)
        temp_zip.writestr(info, data)
assert total == 261126, total
with zipfile.ZipFile(TEMP) as check_zip:
    assert check_zip.testzip() is None
TEMP.replace(TARGET)
print(f'公式缓存恢复完成：{total}', flush=True)
