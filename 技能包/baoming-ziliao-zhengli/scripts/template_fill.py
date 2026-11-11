# -*- coding: utf-8 -*-
"""在模板 docx 原件上做段落级填充，保留模板原有版式（字号/对齐/字体）"""
import os, zipfile, shutil, re
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
for p, u in [('w', W),
             ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
             ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
             ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
             ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')]:
    ET.register_namespace(p, u)


def _w(tag):
    return '{%s}%s' % (W, tag)


def _para_text(p):
    return ''.join(t.text or '' for t in p.iter(_w('t')))


def _set_text(p, text, size=None):
    """保留段落第一个 run 的 rPr，替换文字"""
    runs = p.findall(_w('r'))
    if not runs:
        r = ET.SubElement(p, _w('r'))
        t = ET.SubElement(r, _w('t'))
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return p
    first = runs[0]
    for extra in runs[1:]:
        p.remove(extra)
    for t in first.findall(_w('t')):
        first.remove(t)
    t = ET.SubElement(first, _w('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    if size:
        rpr = first.find(_w('rPr'))
        if rpr is None:
            rpr = ET.SubElement(first, _w('rPr'))
        sz = rpr.find(_w('sz'))
        if sz is None:
            sz = ET.SubElement(rpr, _w('sz'))
        sz.set(_w('val'), str(size))
    return p


def _clone_para(model_p, text):
    """以某段为模板复制一段（继承 pPr 与首个 run 的 rPr）"""
    import copy
    new = copy.deepcopy(model_p)
    _set_text(new, text)
    return new


def read_paragraphs(path):
    """返回 [(text, jc, size_halfpt, bold)] ，供汇总文档复刻模板版式"""
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read('word/document.xml'))
    out = []
    for p in root.iter(_w('p')):
        ppr = p.find(_w('pPr'))
        jc = '-'
        if ppr is not None and ppr.find(_w('jc')) is not None:
            jc = ppr.find(_w('jc')).get(_w('val'))
        size, bold = 24, False
        r = p.find(_w('r'))
        if r is not None:
            rpr = r.find(_w('rPr'))
            if rpr is not None:
                if rpr.find(_w('sz')) is not None:
                    size = int(rpr.find(_w('sz')).get(_w('val')))
                bold = rpr.find(_w('b')) is not None
        out.append({'text': _para_text(p), 'jc': jc, 'size': size, 'bold': bold})
    return out


def _drawing_para(rel_id, pic_id, name, cx, cy, jc='center'):
    """构造一个居中图片段落（inline 图）"""
    frag = (
        '<w:p xmlns:w="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>'
        '<w:jc w:val="%s"/></w:pPr>'
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="Picture %d"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
        % (W, jc, cx, cy, pic_id, pic_id, pic_id, name, rel_id, cx, cy))
    return ET.fromstring(frag)


def fill(src, dst, ops):
    """ops: [('set', idx, text), ('insert_after', idx, [text, ...]), ('set_size', idx, half_pt),
             ('insert_image_after', idx, path, width_cm)]
    返回填充后 docx 路径"""
    zin = zipfile.ZipFile(src)
    root = ET.fromstring(zin.read('word/document.xml'))
    body = root.find(_w('body'))
    paras = [el for el in body if el.tag == _w('p')]

    inserts = {}   # idx -> [new paragraph elements]
    media, rels_add = [], []
    names = zin.namelist()
    n_img = len([x for x in names if re.match(r'word/media/image\d+\.', x)])
    exist_rels = set(re.findall(r'Id="([^"]+)"',
                                zin.read('word/_rels/document.xml.rels').decode('utf8')))
    for op in ops:
        kind = op[0]
        if kind == 'set':
            _, idx, text = op
            _set_text(paras[idx], text)
        elif kind == 'set_size':
            _, idx, size = op
            for r in paras[idx].findall(_w('r')):
                rpr = r.find(_w('rPr'))
                if rpr is None:
                    rpr = ET.SubElement(r, _w('rPr'))
                sz = rpr.find(_w('sz'))
                if sz is None:
                    sz = ET.SubElement(rpr, _w('sz'))
                sz.set(_w('val'), str(size))
        elif kind == 'insert_after':
            _, idx, texts = op
            inserts.setdefault(idx, []).extend(_clone_para(paras[idx], t) for t in texts)
        elif kind == 'insert_image_after':
            _, idx, path, width_cm = op
            from docxgen import img_size_px
            w_px, h_px = img_size_px(path)
            cx = int(width_cm * 360000)
            cy = int(cx * h_px / w_px)
            while True:
                n_img += 1
                rel_id = 'rIdTplImg%d' % n_img
                if rel_id not in exist_rels:
                    break
            ext = 'png' if path.lower().endswith('.png') else 'jpeg'
            part = 'word/media/image%d.%s' % (n_img, ext)
            with open(path, 'rb') as f:
                media.append((part, f.read()))
            rels_add.append((rel_id, 'media/image%d.%s' % (n_img, ext), 'image/%s' % ext))
            inserts.setdefault(idx, []).append(
                _drawing_para(rel_id, 900 + n_img, 'image%d.%s' % (n_img, ext), cx, cy))

    for idx, news in inserts.items():
        pos = list(body).index(paras[idx])
        for k, np in enumerate(news, 1):
            body.insert(pos + k, np)

    xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

    rels_xml, ct_xml = None, None
    if rels_add:
        PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ET.register_namespace('pr', PKG)
        rr = ET.fromstring(zin.read('word/_rels/document.xml.rels'))
        for rid, target, _ct in rels_add:
            e = ET.SubElement(rr, '{%s}Relationship' % PKG)
            e.set('Id', rid)
            e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/'
                           'relationships/image')
            e.set('Target', target)
        rels_xml = ET.tostring(rr, encoding='UTF-8', xml_declaration=True)

        ct = zin.read('[Content_Types].xml').decode('utf8')
        for ext in sorted({os.path.splitext(p)[1].lstrip('.').lower() for p, _ in media}):
            if 'Extension="%s"' % ext not in ct:
                ct = ct.replace('</Types>',
                                '<Default Extension="%s" ContentType="image/%s"/></Types>'
                                % (ext, ext))
        ct_xml = ct.encode('utf8')

    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            if n == 'word/document.xml':
                zout.writestr(n, xml)
            elif n == 'word/_rels/document.xml.rels' and rels_xml:
                zout.writestr(n, rels_xml)
            elif n == '[Content_Types].xml' and ct_xml:
                zout.writestr(n, ct_xml)
            elif n.endswith('/'):
                continue
            else:
                zout.writestr(n, zin.read(n))
        for part, data in media:
            zout.writestr(part, data)
    zin.close()
    return dst
