# -*- coding: utf-8 -*-
"""无第三方依赖的 DOCX 生成器（A4 / 宋体小四 / 标题2 / 表格 / 内嵌图片）"""
import zipfile, os, struct, datetime

EMU_PER_CM = 360000
CONTENT_W_CM = 14.64          # A4 21cm - 左右各 3.18cm
SANS = '宋体'
FANG = '仿宋'
ASCII_FONT = 'Times New Roman'


def _esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def img_size_px(path):
    """返回 (w,h) 像素，支持 png / jpeg"""
    with open(path, 'rb') as f:
        head = f.read(26)
        if head[:8] == b'\x89PNG\r\n\x1a\n':
            w, h = struct.unpack('>II', head[16:24])
            return w, h
        f.seek(0)
        data = f.read()
    if data[:2] != b'\xff\xd8':
        raise ValueError('不支持的图片格式: %s' % path)
    i = 2
    while i < len(data) - 9:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            h, w = struct.unpack('>HH', data[i + 5:i + 9])
            return w, h
        seg = struct.unpack('>H', data[i + 2:i + 4])[0]
        i += 2 + seg
    raise ValueError('无法解析 JPEG 尺寸: %s' % path)


RPR_BODY = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>' % (ASCII_FONT, ASCII_FONT, SANS))
RPR_FANG = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>' % (ASCII_FONT, ASCII_FONT, FANG))


def rpr(font='宋体', size=12, bold=False, color=None):
    """按 OOXML schema 顺序构造 rPr：rFonts → b → color → sz → szCs"""
    return ((RPR_BODY if font == '宋体' else RPR_FANG)
            + ('<w:b/>' if bold else '')
            + ('<w:color w:val="%s"/>' % color if color else '')
            + '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (int(size * 2), int(size * 2)))


class Doc:
    def __init__(self):
        self.body = []
        self.media = []          # (name, bytes)
        self.rels = []
        self._img_idx = 0
        self._bookmark = 0

    # ---------- 基础段落 ----------
    def p(self, text='', size=12, bold=False, font='宋体', align='left',
          indent_first=2, space_before=0, space_after=0, line=1.5, style=None,
          color=None):
        """size 单位 pt；indent_first 首行缩进字符数；color 如 'FF0000'"""
        ppr = ['<w:pPr>']
        if style:
            ppr.append('<w:pStyle w:val="%s"/>' % style)
        ppr.append('<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
                   % (space_before * 20, space_after * 20, int(line * 240)))
        ppr.append('<w:ind w:firstLineChars="%d" w:firstLine="%d"/>'
                   % (indent_first * 100, int(indent_first * size * 20)))
        ppr.append('<w:jc w:val="%s"/>' % align)
        ppr.append('<w:rPr>%s</w:rPr>' % rpr(font, size, bold, color))
        ppr.append('</w:pPr>')
        runs = ''
        if text != '':
            r = rpr(font, size, bold, color)
            for i, seg in enumerate(str(text).split('\n')):
                if i:
                    runs += '<w:r><w:rPr>%s</w:rPr><w:br/></w:r>' % r
                runs += '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (r, _esc(seg))
        self.body.append('<w:p>%s%s</w:p>' % (''.join(ppr), runs))

    def title(self, text, size=18):
        self.p(text, size=size, bold=True, align='center', indent_first=0,
               space_before=0, space_after=10, line=1.5, style='Title')

    def h2(self, text):
        self.p(text, size=15, bold=True, font='宋体', align='left', indent_first=0,
               space_before=10, space_after=6, line=1.5, style='Heading2')

    def p_runs(self, segments, size=12, indent_first=0, align='left',
               space_before=0, space_after=0, line=1.5):
        """segments: [(text, opts)] opts 支持 bold / highlight / font"""
        ppr = ('<w:pPr><w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
               '<w:ind w:firstLineChars="%d" w:firstLine="%d"/><w:jc w:val="%s"/></w:pPr>'
               % (space_before * 20, space_after * 20, int(line * 240),
                  indent_first * 100, int(indent_first * size * 20), align))
        runs = []
        for text, o in segments:
            o = o or {}
            r = rpr(o.get('font', '宋体'), size, o.get('bold', False))
            if o.get('highlight'):
                r += '<w:highlight w:val="%s"/>' % o['highlight']
            runs.append('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                        % (r, _esc(text)))
        self.body.append('<w:p>%s%s</w:p>' % (ppr, ''.join(runs)))

    def toc_line(self, left, right, size=12):
        """目录行：左标题 + 右对齐状态，中间点选线引导"""
        ppr = ('<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8300"/></w:tabs>'
               '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
               '<w:ind w:firstLine="0"/><w:jc w:val="left"/>'
               '<w:rPr>%s</w:rPr></w:pPr>' % rpr('宋体', size, False))
        runs = ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                '<w:r><w:rPr>%s</w:rPr><w:tab/><w:t xml:space="preserve">%s</w:t></w:r>'
                % (rpr('宋体', size, False), _esc(left), rpr('宋体', size, False), _esc(right)))
        self.body.append('<w:p>%s%s</w:p>' % (ppr, runs))

    def blank(self, n=1):
        for _ in range(n):
            self.p('', indent_first=0)

    def page_break(self):
        self.body.append('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
                         '<w:r><w:br w:type="page"/></w:r></w:p>')

    # ---------- 图片 ----------
    def image(self, path, width_cm=None, align='center'):
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        w_px, h_px = img_size_px(path)
        if width_cm is None:
            width_cm = min(CONTENT_W_CM, 16.0)
        if width_cm > CONTENT_W_CM:
            width_cm = CONTENT_W_CM
        width_emu = int(width_cm * EMU_PER_CM)
        height_emu = int(width_emu * h_px / w_px)

        self._img_idx += 1
        idx = self._img_idx
        ext = 'png' if path.lower().endswith('.png') else 'jpeg'
        name = 'image%d.%s' % (idx, ext)
        rel_id = 'rIdImg%d' % idx
        with open(path, 'rb') as f:
            self.media.append((name, f.read()))
        self.rels.append((rel_id, 'media/%s' % name, 'image/%s' % ext))

        drawing = (
            '<w:p><w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="%s"/></w:pPr><w:r><w:drawing>'
            '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:docPr id="%d" name="Picture %d"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
            % (align, width_emu, height_emu, idx, idx, idx, name, rel_id, width_emu, height_emu)
        )
        self.body.append(drawing)

    # ---------- 表格 ----------
    def table(self, header, rows, widths_cm=None, font_size=10.5, align_center_cols=None,
              colors=None):
        """colors: {(行号, 列号): 'FF0000'}，行号从 0 起（数据行，不含表头）"""
        ncol = len(header)
        if widths_cm is None:
            widths_cm = [CONTENT_W_CM / ncol] * ncol
        total = sum(widths_cm)
        widths_cm = [w * CONTENT_W_CM / total for w in widths_cm]
        align_center_cols = align_center_cols or []
        colors = colors or {}

        def cell(txt, w, bold=False, center=False, color=None, shade=False):
            w_tw = int(w * 567)
            return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
                    '<w:vAlign w:val="center"/></w:tcPr>'
                    '<w:p><w:pPr><w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/>'
                    '<w:ind w:firstLine="0"/><w:jc w:val="%s"/></w:pPr>'
                    '<w:r><w:rPr>%s</w:rPr>'
                    '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc>'
                    % (w_tw, '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if shade else '',
                       'center' if center else 'left',
                       rpr('宋体', font_size, bold, color), _esc(txt)))

        borders = ('<w:tblBorders>'
                   + ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % s
                             for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
                   + '</w:tblBorders>')
        xml = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>%s'
               '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
               '<w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
               '</w:tblPr><w:tblGrid>' % (int(sum(widths_cm) * 567), borders)]
        xml.append(''.join('<w:gridCol w:w="%d"/>' % int(w * 567) for w in widths_cm))
        xml.append('</w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>')
        for c, h in zip(widths_cm, header):
            xml.append(cell(h, c, bold=True, center=True, shade=True))
        xml.append('</w:tr>')
        for ri, row in enumerate(rows):
            xml.append('<w:tr>')
            for ci, (c, v) in enumerate(zip(widths_cm, row)):
                xml.append(cell(v, c, center=(ci in align_center_cols),
                                color=colors.get((ri, ci))))
            xml.append('</w:tr>')
        xml.append('</w:tbl>')
        self.body.append(''.join(xml))

    # ---------- 保存 ----------
    def save(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
                'w:header="851" w:footer="992" w:gutter="0"/>'
                '<w:cols w:space="425"/><w:docGrid w:linePitch="312"/></w:sectPr>')
        document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:document '
                    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                    '<w:body>' + ''.join(self.body) + sect + '</w:body></w:document>')

        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                + ''.join('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>' % (i, t)
                          for i, t, _ in self.rels)
                + '</Relationships>')

        img_ct = ''.join('<Default Extension="%s" ContentType="image/%s"/>' % (e.split('/')[1], e)
                         for e in sorted({r[2] for r in self.rels}))
        imgs = ''.join('<Override PartName="/word/media/%s" ContentType="%s"/>' % (n, e)
                       for (n, _), (_, _, e) in zip(self.media, self.rels))
        ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              + img_ct +
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
              '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
              '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
              '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
              + imgs + '</Types>')

        root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                     '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                     '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                     '</Relationships>')

        now = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
        core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                '<dc:title>报名资料汇总</dc:title><dc:creator>workbuddy</dc:creator>'
                '<cp:lastModifiedBy>workbuddy</cp:lastModifiedBy>'
                '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                '</cp:coreProperties>' % (now, now))
        app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
               'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
               '<Application>Microsoft Office Word</Application><Company></Company></Properties>')

        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                  '<w:docDefaults><w:rPrDefault><w:rPr>' + RPR_BODY +
                  '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
                  '<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                  '</w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/>'
                  '<w:qFormat/><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
                  '<w:rPr>' + RPR_BODY + '<w:sz w:val="24"/></w:rPr></w:style>'
                  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
                  '<w:basedOn w:val="1"/><w:qFormat/><w:pPr><w:jc w:val="center"/></w:pPr>'
                  '<w:rPr>' + RPR_BODY + '<w:b/><w:sz w:val="36"/></w:rPr></w:style>'
                  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
                  '<w:basedOn w:val="1"/><w:next w:val="1"/><w:qFormat/><w:pPr><w:keepNext/>'
                  '<w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>'
                  '<w:rPr>' + RPR_BODY + '<w:b/><w:sz w:val="30"/></w:rPr></w:style>'
                  '</w:styles>')

        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', ct)
            z.writestr('_rels/.rels', root_rels)
            z.writestr('docProps/core.xml', core)
            z.writestr('docProps/app.xml', app)
            z.writestr('word/document.xml', document)
            z.writestr('word/styles.xml', styles)
            z.writestr('word/_rels/document.xml.rels', rels)
            for name, data in self.media:
                z.writestr('word/media/' + name, data)
        return path
