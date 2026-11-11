# -*- coding: utf-8 -*-
"""无第三方依赖的 xlsx / docx 读取工具"""
import sys, zipfile, re
import xml.etree.ElementTree as ET

NS = {
 'm':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
 'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
 'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
 'pr':'http://schemas.openxmlformats.org/package/2006/relationships',
}

def col2idx(ref):
    m = re.match(r'([A-Z]+)', ref)
    if not m: return 0
    s = m.group(1); n = 0
    for ch in s: n = n*26 + (ord(ch)-64)
    return n-1

def read_xlsx(path):
    z = zipfile.ZipFile(path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall('m:si', NS):
            txt = ''.join(t.text or '' for t in si.iter('{%s}t' % NS['m']))
            shared.append(txt)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap = {r.get('Id'): r.get('Target') for r in rels}
    out = []
    for sh in wb.find('m:sheets', NS):
        name = sh.get('name')
        rid = sh.get('{%s}id' % NS['r'])
        tgt = rmap[rid]
        if not tgt.startswith('xl/'): tgt = 'xl/' + tgt.lstrip('/')
        ws = ET.fromstring(z.read(tgt))
        rows = []
        for row in ws.iter('{%s}row' % NS['m']):
            cells = {}
            for c in row.findall('m:c', NS):
                ref = c.get('r'); ci = col2idx(ref)
                t = c.get('t')
                v = c.find('m:v', NS)
                isn = c.find('m:is', NS)
                if t == 's' and v is not None:
                    val = shared[int(v.text)]
                elif t == 'inlineStr' and isn is not None:
                    val = ''.join(x.text or '' for x in isn.iter('{%s}t' % NS['m']))
                elif v is not None:
                    val = v.text
                else:
                    val = ''
                cells[ci] = (val or '').strip()
            if cells:
                n = max(cells)+1
                rows.append([cells.get(i,'') for i in range(n)])
        out.append((name, rows))
    return out

def read_docx(path):
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read('word/document.xml'))
    W = '{%s}' % NS['w']
    lines = []
    body = root.find('w:body', NS)
    for el in body:
        tag = el.tag.replace(W,'')
        if tag == 'p':
            t = ''.join(x.text or '' for x in el.iter(W+'t'))
            style = el.find('w:pPr/w:pStyle', NS)
            st = style.get(W+'val') if style is not None else ''
            pics = len(list(el.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip')))
            lines.append(('[%s]' % st if st else '') + t + (' <图片x%d>' % pics if pics else ''))
        elif tag == 'tbl':
            lines.append('<表格>')
            for tr in el.findall('w:tr', NS):
                cells = []
                for tc in tr.findall('w:tc', NS):
                    cells.append(''.join(x.text or '' for x in tc.iter(W+'t')))
                lines.append(' | '.join(cells))
            lines.append('</表格>')
    return '\n'.join(lines), z

if __name__ == '__main__':
    p = sys.argv[1]
    if p.lower().endswith('.xlsx'):
        for name, rows in read_xlsx(p):
            print('=== SHEET:', name, len(rows))
            for i, r in enumerate(rows, 1):
                print(i, '|', ' || '.join(r))
    else:
        txt, _ = read_docx(p)
        print(txt)
