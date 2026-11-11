# -*- coding: utf-8 -*-
"""把匹配到的资质附件插入甲方 Word 入围文档（生成副本，不修改原件）。

用法:
    python build_filed_word.py --docx <甲方Word文件> --match <match_result.json> \
        --out <输出文件.docx> [--appendix] \
        [--ledgers <parse_ledgers.py 输出目录>]

插入规则:
    --inline（默认）: 在文档中定位每条需求的段落，在其下方插入
        『资质名称』标题 + 证书原件（jpg/png 直接插图；PDF 等非图片以超链接附于该处）。
        文档中找不到对应段落的条目 → 统一插入文档末尾并记录。
    --appendix: 全部附件统一附在 Word 文档末尾，按甲方文档列举顺序排列，
        每项上方以资质名称作为标题。
    文档末尾强制追加提示: 【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】

硬性规则: 源文件只读，输出为副本；无法插入的位置如实记录，禁止虚构成功。
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import MARK_OK, MARK_MANUAL, MARK_MISSING, die  # noqa: E402

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
except ImportError:
    die("缺少 python-docx，请安装: pip install python-docx")

FOOTER_NOTE = "【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}


NOISE = re.compile(r"[，。、；：,.;:（）()\s　「」『』“”\"'【】\[\]一二三四五六七八九十0-9]+")


def clean(s):
    return NOISE.sub("", str(s))


def add_hyperlink(paragraph, url, text):
    """在段落中插入可点击超链接（指向本地副本文件）。"""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = paragraph._p.makeelement(qn("w:hyperlink"), {qn("r:id"): r_id})
    new_run = paragraph._p.makeelement(qn("w:r"), {})
    rPr = paragraph._p.makeelement(qn("w:rPr"), {})
    rStyle = paragraph._p.makeelement(qn("w:rStyle"), {qn("w:val"): "Hyperlink"})
    rPr.append(rStyle)
    new_run.append(rPr)
    t = paragraph._p.makeelement(qn("w:t"), {})
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def find_para_index(doc, req_text):
    """在文档中查找包含需求文本（或核心关键词）的段落，返回索引；找不到返回 None。"""
    rc = clean(req_text)
    if not rc:
        return None
    # 用需求文本最长的片段做匹配，避免过短误配
    for i, para in enumerate(doc.paragraphs):
        pc = clean(para.text)
        if pc and (rc in pc or pc in rc):
            return i
    return None


def insert_attachment(doc, para, title, src_path, copy_dir):
    """在段落 para 之后插入: 标题段 + 附件（图片直接插 / 其他格式超链接）。
    返回 (插入说明, 附件副本文件名)。"""
    # 标题
    h = doc.add_paragraph()
    run = h.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # 附件
    ext = os.path.splitext(src_path)[1].lower()
    if ext in IMAGE_EXTS:
        p = doc.add_paragraph()
        p.add_run().add_picture(src_path, width=Inches(5.5))
        note = f"图片已插入: {os.path.basename(src_path)}"
    else:
        p = doc.add_paragraph()
        p.add_run("附件原件（非图片格式，点击链接打开本地副本）: ")
        add_hyperlink(p, src_path, os.path.basename(src_path))
        note = f"非图片附件以超链接附上: {os.path.basename(src_path)}"
    # 把新增段落移动到 para 之后（inline 模式）
    if para is not None:
        for new_p in (h, p):
            para._p.addnext(new_p._p)
        # 调整顺序：h 应紧邻 para 之后，p 在 h 之后
        h._p.addnext(p._p)
    return note, os.path.basename(src_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True, help="甲方 Word 入围文档")
    ap.add_argument("--match", required=True, help="match_result.json")
    ap.add_argument("--out", required=True, help="输出 docx 路径")
    ap.add_argument("--appendix", action="store_true", help="附件统一附在末尾")
    ap.add_argument("--ledgers", default=None)
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        die(f"甲方 Word 文档不存在: {args.docx}")
    if not os.path.isfile(args.match):
        die(f"匹配结果不存在: {args.match}")

    results = json.load(open(args.match, encoding="utf-8"))
    attached = [r for r in results if r.get("附件源路径") and os.path.isfile(r.get("附件源路径", ""))]
    if not attached:
        die("匹配结果中没有可插入的附件，请先核对 match_result.json（禁止虚构附件）")

    try:
        doc = Document(args.docx)
    except Exception as e:
        die(f"Word 文档读取失败: {e}")

    log = []
    appendix_items = []  # (title, src) 需要附到末尾的

    if not args.appendix:
        # inline: 逐条定位
        for r in attached:
            idx = find_para_index(doc, r.get("甲方需求", ""))
            if idx is None:
                appendix_items.append((r.get("甲方需求", ""), r["附件源路径"]))
                log.append(f"⚠️ 未定位到对应段落，改附文档末尾: {r.get('甲方需求','')}")
            else:
                para = doc.paragraphs[idx]
                title = r.get("甲方需求", "")
                note, fname = insert_attachment(doc, para, title, r["附件源路径"], None)
                log.append(f"✅ 已插入『{title}』于段落「{para.text[:30]}」之后: {note}")
    else:
        for r in attached:
            appendix_items.append((r.get("甲方需求", ""), r["附件源路径"]))

    # 末尾附件
    for title, src in appendix_items:
        h = doc.add_paragraph()
        run = h.add_run(title)
        run.bold = True
        run.font.size = Pt(12)
        ext = os.path.splitext(src)[1].lower()
        if ext in IMAGE_EXTS:
            p = doc.add_paragraph()
            p.add_run().add_picture(src, width=Inches(5.5))
        else:
            p = doc.add_paragraph()
            p.add_run("附件原件: ")
            add_hyperlink(p, src, os.path.basename(src))
        log.append(f"📎 文档末尾附件: {title} → {os.path.basename(src)}")

    # 强制末尾提示
    fn = doc.add_paragraph()
    fn_run = fn.add_run(FOOTER_NOTE)
    fn_run.bold = True
    fn_run.font.color.rgb = None
    fn_run.font.size = Pt(10)

    out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    doc.save(out)

    report = ["# Word 资质附件插入报告", ""]
    report.append(f"- 输入: {args.docx}")
    report.append(f"- 输出副本: {out}")
    report.append(f"- 模式: {'统一附末尾' if args.appendix else '对应条目下方(定位失败者附末尾)'}")
    report += log
    report.append("")
    report.append("> 【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】")
    report_path = os.path.join(os.path.dirname(out), "word_insert_report.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report))

    print(f"✅ Word 处理完成: {out}")
    print(f"   报告: {report_path}")
    for ln in log:
        print("   " + ln)
    print("⚠️ 请人工打开成品文档复核：图片位置、超链接可用性、与甲方要求的一致性。")


if __name__ == "__main__":
    main()
