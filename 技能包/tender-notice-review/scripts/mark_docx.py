# -*- coding: utf-8 -*-
"""
对招标公告 Word(.docx) 副本施加高亮底纹 + 批注。
硬性约束：只读取输入文件，输出到新的副本路径，绝不修改原文件。

用法：
  python mark_docx.py --tasks 任务.json
任务 JSON 格式：
{
  "input":  "原公告路径.docx",
  "output": "标记版副本路径.docx",
  "tasks": [
    {"search": "公告原文中精确存在的文本片段", "color": "yellow|orange|green",
     "comment": "批注内容", "occurrence": 0}
  ]
}
说明：
  - search 必须是公告原文中的精确子串（AI 先读取原文提取，禁止凭空编写）；
  - occurrence 可选，从0开始，用于同一段文本多次出现时只标第N次；缺省标全部出现；
  - 颜色：yellow=时间类，orange=特殊报名方式/原件要求，green=CA/重点平台命中；
  - 未匹配到的任务会在结果中列出，供 AI 修正后重跑，禁止假装标记成功。
"""
import argparse
import copy
import json
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

COLORS = {
    "yellow": "FFFF00",   # 时间类
    "orange": "FFC000",   # 特殊报名方式 / 原件要求
    "green": "C6EFCE",    # CA / 重点平台命中
}


def set_run_shading(run, fill):
    rpr = run._r.get_or_add_rPr()
    shd = rpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        rpr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def runs_with_offsets(paragraph):
    out = []
    c = 0
    for r in paragraph.runs:
        n = len(r.text)
        out.append((r, c, c + n))
        c += n
    return out


def split_run(run, offset):
    """把 run 在其文本的第 offset 字符处切分为两个 run，返回(左, 右)。保留原格式。"""
    from docx.text.run import Run
    new_el = copy.deepcopy(run._r)
    run._r.addnext(new_el)
    right = Run(new_el, run._parent)
    right.text = run.text[offset:]
    run.text = run.text[:offset]
    return run, right


def split_at(paragraph, pos):
    """在段落文本的第 pos 字符处切分跨界的 run。发生切分返回 True。"""
    for r, rs, re_ in runs_with_offsets(paragraph):
        if rs < pos < re_ and len(r.text) > 1:
            split_run(r, pos - rs)
            return True
    return False


def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def mark_paragraph(doc, paragraph, search, color, comment, occurrence=None):
    """在段落内标记 search 的出现位置。返回本段落标记次数。"""
    count = 0
    full = paragraph.text
    start = 0
    while True:
        s = full.find(search, start)
        if s < 0:
            break
        e = s + len(search)
        if occurrence is None or count == occurrence:
            # 先切分跨界 run（多次直到稳定）
            for _ in range(8):
                if not split_at(paragraph, s):
                    break
            for _ in range(8):
                if not split_at(paragraph, e):
                    break
            matched = []
            c = 0
            for r in paragraph.runs:
                n = len(r.text)
                if n and c >= s and c + n <= e:
                    matched.append(r)
                c += n
            fill = COLORS.get(color, COLORS["yellow"])
            for r in matched:
                set_run_shading(r, fill)
            if comment and matched:
                try:
                    doc.add_comment(
                        matched,
                        text=comment,
                        author="AI辅助分析(须人工复核)",
                        initials="AI",
                    )
                except Exception as ex:
                    print(f"警告：添加批注失败（{ex}），高亮已生效：{search[:30]}", file=sys.stderr)
        count += 1
        start = e
        if occurrence is not None and count > occurrence:
            break
    return count


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", required=True, help="任务JSON文件路径")
    args = ap.parse_args()

    with open(args.tasks, "r", encoding="utf-8") as f:
        spec = json.load(f)

    doc = Document(spec["input"])
    report = {"output": spec["output"], "matched": [], "unmatched": []}

    for task in spec.get("tasks", []):
        search = task.get("search", "")
        if not search:
            continue
        color = task.get("color", "yellow")
        comment = task.get("comment", "")
        occurrence = task.get("occurrence")
        total = 0
        for p in iter_paragraphs(doc):
            if search in p.text:
                total += mark_paragraph(doc, p, search, color, comment, occurrence)
        if total > 0:
            report["matched"].append({"search": search, "hits": total, "color": color})
        else:
            report["unmatched"].append({"search": search, "color": color,
                                        "reason": "公告原文中未找到该精确文本，需重新提取原文片段"})

    doc.save(spec["output"])
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["unmatched"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
