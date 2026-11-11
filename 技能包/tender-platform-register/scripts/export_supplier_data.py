# -*- coding: utf-8 -*-
"""读取供应商资料库三张核心表，输出结构化缓存 JSON。

只读原始 xlsx，不修改任何文件。仅依赖标准库（zipfile + xml.etree），
不依赖 openpyxl，也不需要 Excel 编辑器。

用法：
    python export_supplier_data.py --root "E:\\卫华供应商入围资料库" --out supplier_cache.json

输出结构：
    base_info          基础信息表各子表的键值对
    qualifications     资质清单条目（含 exists / abs_path）
    documents          入围资料清单条目（含 exists / corrected_path）
    missing            缺失文件汇总
    upload_candidates  按常见页面控件名归类的候选文件
"""
import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

BASE_XLSX = '供方能力调查基础信息表.xlsx'
QUAL_XLSX = '资质清单.xlsx'
DOCS_XLSX = '供应商入围资料清单.xlsx'


def _col_to_idx(letters):
    c = 0
    for ch in letters:
        c = c * 26 + (ord(ch) - 64)
    return c


def read_workbook(path):
    """返回 {sheet_name: [[cell,...], ...]}，行列均为字符串，缺失为 ''。"""
    if not os.path.exists(path):
        raise SystemExit('【文件不存在：%s】' % path)
    z = zipfile.ZipFile(path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(NS + 'si'):
            shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    out = {}
    for sh in wb.find(NS + 'sheets'):
        name = sh.get('name')
        tgt = relmap.get(sh.get(RNS + 'id'), '')
        if tgt.startswith('/'):
            tgt = tgt[1:]
        elif not tgt.startswith('xl/'):
            tgt = 'xl/' + tgt
        if tgt not in z.namelist():
            continue
        root = ET.fromstring(z.read(tgt))
        sd = root.find(NS + 'sheetData')
        rows = []
        if sd is None:
            out[name] = rows
            continue
        for row in sd.findall(NS + 'row'):
            cells = {}
            for c in row.findall(NS + 'c'):
                ref = c.get('r') or ''
                m = re.match(r'([A-Z]+)(\d+)', ref)
                if not m:
                    continue
                ci = _col_to_idx(m.group(1))
                t = c.get('t')
                v = c.find(NS + 'v')
                isel = c.find(NS + 'is')
                val = ''
                if t == 's' and v is not None:
                    try:
                        val = shared[int(v.text)]
                    except (ValueError, IndexError):
                        val = v.text or ''
                elif t == 'inlineStr' and isel is not None:
                    val = ''.join(x.text or '' for x in isel.iter(NS + 't'))
                elif v is not None:
                    val = v.text or ''
                cells[ci] = str(val).replace('\n', ' / ').strip()
            if cells:
                width = max(cells)
                rows.append([cells.get(i, '') for i in range(1, width + 1)])
        out[name] = rows
    return out


def kv_from_sheet(rows):
    """两列表格 → 有序键值对。跳过段落标题（只有 A 列有值的行）。"""
    data = []
    for r in rows:
        a = r[0] if len(r) > 0 else ''
        b = r[1] if len(r) > 1 else ''
        if not b:
            continue
        if a.startswith('项目') or a.startswith('字段'):
            continue
        data.append({'field': a, 'value': b})
    return data


def table_rows(rows, header_aliases):
    """按表头名称定位列，返回字典列表。"""
    if not rows:
        return []
    header = rows[0]
    idx = {}
    for key, aliases in header_aliases.items():
        for i, h in enumerate(header):
            for alias in aliases:
                if alias in (h or ''):
                    idx[key] = i
                    break
            if key in idx:
                break
    out = []
    for r in rows[1:]:
        item = {}
        for key, i in idx.items():
            item[key] = r[i] if i < len(r) else ''
        if any(item.values()):
            out.append(item)
    return out


def normalize_rel(p):
    return p.replace('/', os.sep).replace('\\', os.sep).strip()


_PUNCT = re.compile(r'[\s\u3000\.\-_（）()【】\[\]、,，。;；:：+\'"“”‘’·|]')


def fuzzy_key(name):
    """去掉空白与标点、统一小写，用于路径失效后的模糊召回。"""
    base = os.path.splitext(name)[0]
    return _PUNCT.sub('', base).lower()


def build_index(root):
    """全库文件名索引：精确名 -> 路径，模糊键 -> 路径列表。"""
    exact, fuzzy = {}, {}
    for dirpath, _dirnames, filenames in os.walk(root):
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            exact.setdefault(fn, []).append(full)
            fuzzy.setdefault(fuzzy_key(fn), []).append(full)
    return exact, fuzzy


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--root', default=r'E:\卫华供应商入围资料库')
    ap.add_argument('--out', default='supplier_cache.json')
    args = ap.parse_args()

    root = args.root
    if not os.path.isdir(root):
        raise SystemExit('【文件不存在：%s】' % root)

    # 1) 基础信息
    wb = read_workbook(os.path.join(root, BASE_XLSX))
    sheets = list(wb.keys())
    base_sheet = sheets[0]
    base_info = {
        'sheet': base_sheet,
        'sheets_all': sheets,
        'items': kv_from_sheet(wb[base_sheet]),
    }

    # 2) 资质清单
    wbq = read_workbook(os.path.join(root, QUAL_XLSX))
    q_rows = table_rows(list(wbq.values())[0], {
        'category': ['类别'],
        'cert_name': ['证书完整名称', '名称'],
        'path': ['文件路径', '路径'],
        'status': ['状态'],
        'note': ['备注'],
        'validity': ['有效期'],
    })
    qualifications = []
    for it in q_rows:
        rel = normalize_rel(it.get('path', ''))
        abs_path = os.path.join(root, rel) if rel else ''
        qualifications.append({
            'category': it.get('category', ''),
            'cert_name': it.get('cert_name', ''),
            'rel_path': rel,
            'abs_path': abs_path,
            'exists': bool(abs_path) and os.path.exists(abs_path),
            'status': it.get('status', ''),
            'validity': it.get('validity', ''),
            'note': it.get('note', ''),
        })

    # 3) 入围资料清单
    wbd = read_workbook(os.path.join(root, DOCS_XLSX))
    d_rows = table_rows(list(wbd.values())[0], {
        'name': ['文件名'],
        'path': ['文件完整路径', '路径'],
        'ext': ['文件后缀', '后缀'],
        'size_kb': ['文件大小', '大小'],
        'note': ['备注'],
    })
    documents = []
    missing = []
    exact_idx, fuzzy_idx = build_index(root)
    for it in d_rows:
        name = it.get('name', '')
        raw = it.get('path', '').strip()
        exists = bool(raw) and os.path.exists(raw)
        corrected = ''
        match_mode = 'exact' if exists else ''
        if not exists and name:
            d = os.path.dirname(raw) if raw else ''
            guess = os.path.join(d, name) if d else ''
            if guess and os.path.exists(guess):
                corrected = guess
                match_mode = 'same_dir_name'
            elif exact_idx.get(name):
                # 同名文件被移动到资料库其他位置
                corrected = exact_idx[name][0]
                match_mode = 'moved_same_name'
            else:
                hits = fuzzy_idx.get(fuzzy_key(name), [])
                if len(hits) == 1:
                    corrected = hits[0]
                    match_mode = 'fuzzy_unique'
                elif len(hits) > 1:
                    corrected = hits[0]
                    match_mode = 'fuzzy_ambiguous'
        if corrected:
            exists = True
        ext_changed = bool(corrected) and (
            os.path.splitext(name)[1].lower() != os.path.splitext(corrected)[1].lower())
        item = {
            'name': name,
            'raw_path': raw,
            'path': corrected or raw,
            'ext': it.get('ext', ''),
            'size_kb': it.get('size_kb', ''),
            'note': it.get('note', ''),
            'exists': exists,
            'corrected': bool(corrected),
            'match_mode': match_mode,
            'ext_changed': ext_changed,
            'format_note': ('清单登记为 %s，实际文件为 %s，上传前确认页面 accept 是否允许'
                            % (os.path.splitext(name)[1].lstrip('.'), os.path.splitext(corrected)[1].lstrip('.'))
                            if ext_changed else ''),
            'candidates': (fuzzy_idx.get(fuzzy_key(name), [])[:5] if match_mode.startswith('fuzzy') else []),
        }
        documents.append(item)
        if not exists:
            missing.append({'source': 'documents', 'name': name, 'raw_path': raw})

    for q in qualifications:
        if not q['exists']:
            missing.append({'source': 'qualifications', 'name': q['cert_name'], 'raw_path': q['rel_path']})

    # 4) 常见控件名 → 候选文件（按关键词打分，取最高分一条）
    CONTROL_KEYWORDS = [
        ('营业执照', ['营业执照']),
        ('特种设备生产许可证', ['特种设备生产许可证', '特种设备']),
        ('安全生产许可证', ['安全生产许可证']),
        ('建筑业企业资质', ['建筑企业资质', '建筑业企业资质', '建筑资质']),
        ('质量管理体系', ['质量管理体系', 'ISO9001', '9001']),
        ('环境管理体系', ['环境管理体系', 'ISO14001', '14001']),
        ('职业健康安全体系', ['职业健康管理体系', '职业健康', '45001']),
        ('高新技术企业证书', ['高新技术企业']),
        ('开户许可', ['开户许可', '基本存款账户']),
        ('一般纳税人证明', ['一般纳税人']),
        ('信用截图', ['无严重失信', '信用中国', '失信', '黑名单', '行政处罚']),
        ('财务报表', ['财务报表', '财务审计报告']),
        ('业绩合同', ['业绩合同', '业绩清单']),
        ('承诺书', ['承诺书', '承诺']),
        ('法人身份证明', ['法人身份证', '法定代表人身份证', '身份证明']),
        ('授权委托书', ['授权委托书', '授权书']),
        ('厂区照片', ['生产现场照片', '现场环境照片', '厂区']),
    ]
    pool = []
    for q in qualifications:
        if q['exists']:
            pool.append({'name': q['cert_name'], 'path': q['abs_path'], 'from': 'qualifications'})
    for d in documents:
        if d['exists']:
            pool.append({'name': d['name'], 'path': d['path'], 'from': 'documents'})

    upload_candidates = {}
    for control, kws in CONTROL_KEYWORDS:
        hits = [p for p in pool if any(k.lower() in (p['name'] + ' ' + p['path']).lower() for k in kws)]
        upload_candidates[control] = hits[:5]

    result = {
        'root': root,
        'base_info': base_info,
        'qualifications': qualifications,
        'documents': documents,
        'missing': missing,
        'upload_candidates': upload_candidates,
        'stats': {
            'qualifications': len(qualifications),
            'documents': len(documents),
            'missing': len(missing),
        },
    }

    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print('root: %s' % root)
    print('sheets: %s' % ', '.join(sheets))
    print('基础信息条目: %d' % len(base_info['items']))
    print('资质清单: %d 条，缺失 %d' % (
        len(qualifications), sum(1 for q in qualifications if not q['exists'])))
    print('入围资料清单: %d 条，缺失 %d' % (
        len(documents), sum(1 for d in documents if not d['exists'])))
    print('路径补正: %d 条' % sum(1 for d in documents if d['corrected']))
    for mode in ('same_dir_name', 'moved_same_name', 'fuzzy_unique', 'fuzzy_ambiguous'):
        n = sum(1 for d in documents if d.get('match_mode') == mode)
        if n:
            print('  %-16s %d 条' % (mode, n))
    fmt = sum(1 for d in documents if d.get('ext_changed'))
    if fmt:
        print('扩展名变化: %d 条（上传前核对页面 accept，图片类控件不接受 PDF）' % fmt)
    if missing:
        print('--- 缺失明细 ---')
        for m in missing:
            print('  [%s] %s | %s' % (m['source'], m['name'], m['raw_path']))
    print('输出: %s' % os.path.abspath(args.out))


if __name__ == '__main__':
    sys.exit(main())
