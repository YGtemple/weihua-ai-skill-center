#!/usr/bin/env python3
"""Generate a Dafang-style DOCX from validated meeting-minutes JSON."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape

try:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError as exc:  # pragma: no cover - 环境依赖提示
    sys.stderr.write(
        "缺少 python-docx 依赖。请先在当前 Python 环境安装：\n"
        "    pip install -r scripts/requirements.txt\n"
        f"原始导入错误：{exc}\n"
    )
    raise SystemExit(2) from exc

from check_minutes_json import validate

BODY_SIZE = 12
BODY_FONT = "宋体"
WESTERN_FONT = "Times New Roman"

# 跟踪项：原文明确标注“（跟踪项）”时，仅该标记自动加粗，其他文字保持常规。
TRACKING_RE = re.compile(r"(（跟踪项）|\(跟踪项\))")


def set_run_font(run, size=BODY_SIZE, bold=None):
    run.font.name = WESTERN_FONT
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:ascii"), WESTERN_FONT)
    fonts.set(qn("w:hAnsi"), WESTERN_FONT)
    fonts.set(qn("w:eastAsia"), BODY_FONT)
    fonts.set(qn("w:cs"), WESTERN_FONT)


def set_paragraph_format(paragraph, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_indent=0, left_indent=0, before=0, after=0,
                         keep_with_next=False):
    paragraph.alignment = align
    fmt = paragraph.paragraph_format
    fmt.line_spacing = 1.5
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.first_line_indent = Cm(first_indent)
    fmt.left_indent = Cm(left_indent)
    fmt.right_indent = Cm(0)
    fmt.keep_with_next = keep_with_next
    fmt.widow_control = True


def remove_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "nil")


def set_cell_bottom_border(cell, size=8):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:color"), "000000")
    borders.append(bottom)


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    tc_pr.append(margins)


def add_watermark(section, text):
    if not text:
        return
    header = section.header
    paragraph = header.paragraphs[0]
    safe_text = escape(text)
    xml = f'''
    <w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:o="urn:schemas-microsoft-com:office:office">
      <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
        <v:formulas>
          <v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>
          <v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>
          <v:f eqn="if @0 21600 @4"/><v:f eqn="if @0 0 @5"/><v:f eqn="if @0 21600 @6"/>
          <v:f eqn="mid @7 @8"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @6 @7"/>
        </v:formulas>
        <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
        <v:textpath on="t" fitshape="t"/>
        <v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>
        <o:lock v:ext="edit" text="t" shapetype="t"/>
      </v:shapetype>
      <v:shape id="DafangWaterMark" o:spid="_x0000_s2049" type="#_x0000_t136"
        style="position:absolute;margin-left:0;margin-top:0;width:460pt;height:90pt;rotation:315;z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
        fillcolor="#D9D9D9" stroked="f">
        <v:fill opacity="0.38"/>
        <v:textpath style="font-family:'宋体';font-size:44pt" string="{safe_text}"/>
      </v:shape>
    </w:pict>'''
    run = paragraph.add_run()
    run._r.append(parse_xml(xml))


def add_page_number(section):
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    set_run_font(run, size=9)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    value = OxmlElement("w:t")
    value.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instr, separate, value, end):
        run._r.append(node)


def add_title(document, logo_path):
    table = document.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    remove_table_borders(table)
    widths = (Cm(2.0), Cm(14.0), Cm(2.0))
    for cell, width in zip(table.rows[0].cells, widths):
        cell.width = width
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
    left = table.cell(0, 0).paragraphs[0]
    left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if logo_path and logo_path.is_file():
        left.add_run().add_picture(str(logo_path), width=Cm(1.3))
    center = table.cell(0, 1).paragraphs[0]
    center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    center.paragraph_format.space_before = Pt(4)
    center.paragraph_format.space_after = Pt(4)
    set_run_font(center.add_run("会 议 纪 要"), size=42, bold=False)


def add_company_line(document, company, issuer):
    table = document.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    remove_table_borders(table)
    cells = table.rows[0].cells
    for cell, width in zip(cells, (Cm(12.0), Cm(6.0))):
        cell.width = width
        set_cell_bottom_border(cell)
        set_cell_margins(cell, bottom=60)
    p = cells[0].paragraphs[0]
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_run_font(p.add_run(company), size=10.5)
    p = cells[1].paragraphs[0]
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
    set_run_font(p.add_run("签发人："), size=10.5, bold=True)
    set_run_font(p.add_run(issuer), size=10.5)


def add_labeled_paragraph(document, label, value):
    p = document.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_run_font(p.add_run(f"{label}："), bold=True)
    set_run_font(p.add_run(value))
    return p


def add_heading(document, text, level=1):
    """三级标题层级：
    level=1 章节（一、二、三）：12pt 加粗，段前 8pt
    level=2 h2/department/reply：12pt 加粗，段前 3pt
    level=3 h3（1.1/1.2/4.1 等）：11pt 加粗，左缩进 0.5cm，段前 2pt
    """
    p = document.add_paragraph()
    if level == 1:
        before, size, indent = 8, BODY_SIZE, 0
    elif level == 3:
        before, size, indent = 2, 11, 0.5
    else:
        before, size, indent = 3, BODY_SIZE, 0
    set_paragraph_format(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        before=before,
        left_indent=indent,
        keep_with_next=True,
    )
    set_run_font(p.add_run(text), size=size, bold=True)
    return p


def add_body_with_tracking(paragraph, text, default_bold=None):
    """写入正文片段，自动把“（跟踪项）”加粗，其余文字保持 default_bold。"""
    parts = TRACKING_RE.split(text)
    for part in parts:
        if not part:
            continue
        is_tracking = bool(TRACKING_RE.fullmatch(part))
        if default_bold is not None:
            set_run_font(paragraph.add_run(part), bold=default_bold)
        elif is_tracking:
            set_run_font(paragraph.add_run(part), bold=True)
        else:
            set_run_font(paragraph.add_run(part))


def add_body(document, block):
    btype = block.get("type")
    if btype == "page_break":
        document.add_page_break()
        return
    text = block["text"]
    if btype == "h2" or btype == "department" or btype == "reply":
        add_heading(document, text, level=2)
        return
    if btype == "h3":
        add_heading(document, text, level=3)
        return
    if btype == "item":
        level = block.get("level", 0)
        p = document.add_paragraph()
        set_paragraph_format(p, left_indent=0.74 * level)
        add_body_with_tracking(p, text, default_bold=block.get("bold"))
        return
    p = document.add_paragraph()
    set_paragraph_format(p, first_indent=0.85)
    lead = block.get("lead")
    if lead and text.startswith(lead):
        set_run_font(p.add_run(lead), bold=True)
        add_body_with_tracking(p, text[len(lead):], default_bold=block.get("bold"))
    else:
        add_body_with_tracking(p, text, default_bold=block.get("bold"))


def generate(data: dict, output: Path, logo_path: Path | None):
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    normal = document.styles["Normal"]
    normal.font.name = WESTERN_FONT
    normal.font.size = Pt(BODY_SIZE)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(0)

    document.core_properties.title = next(
        (f["value"] for f in data["fields"] if f.get("label") == "会议主题"),
        "卫华会议纪要",
    )
    add_watermark(section, data.get("watermark", "仅供内部浏览，严禁外传"))
    add_page_number(section)
    add_title(document, logo_path)
    add_company_line(document, data["company"], data["issuer"])

    for field in data["fields"]:
        add_labeled_paragraph(document, field["label"], field["value"])

    agenda = data.get("agenda", [])
    if agenda:
        add_heading(document, "会议整体议程")
        for i, item in enumerate(agenda, 1):
            p = document.add_paragraph()
            set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT)
            set_run_font(p.add_run(f"{i}、{item}"), bold=True)

    for section_data in data["sections"]:
        add_heading(document, section_data["heading"])
        for block in section_data.get("blocks", []):
            add_body(document, block)

    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("json_file", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--logo", type=Path, default=Path(__file__).resolve().parent.parent / "assets" / "logo.png")
    parser.add_argument("--allow-placeholders", action="store_true", help="仅用于测试含××占位的模板")
    args = parser.parse_args()
    try:
        # utf-8-sig 自动兼容带 BOM 的 JSON。
        data = json.loads(args.json_file.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"读取JSON失败：{exc}", file=sys.stderr)
        return 2
    errors, _ = validate(data, args.allow_placeholders)
    if errors:
        print("JSON校验失败，未生成DOCX：", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    try:
        generate(data, args.output, args.logo)
    except Exception as exc:
        print(f"生成DOCX失败：{exc}", file=sys.stderr)
        return 1
    print(f"已生成：{args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
