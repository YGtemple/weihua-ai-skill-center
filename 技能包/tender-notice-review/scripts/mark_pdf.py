# -*- coding: utf-8 -*-
"""
对招标公告 PDF 副本施加高亮标注 + 文字批注（弹窗注释）。
硬性约束：只读取输入文件，输出到新的副本路径，绝不修改原文件。

用法：
  python mark_pdf.py --tasks 任务.json
任务 JSON 格式与 mark_docx.py 完全一致：
{
  "input":  "原公告路径.pdf",
  "output": "标记版副本路径.pdf",
  "tasks": [
    {"search": "公告原文中精确存在的文本片段", "color": "yellow|orange|green",
     "comment": "批注内容", "occurrence": 0}
  ]
}
说明：
  - search 必须是 PDF 原文中的精确子串；
  - 每处命中添加高亮 + 弹窗批注，批注以"AI辅助分析(须人工复核)"为作者标识；
  - 未匹配任务会在结果中列出，禁止假装标记成功。
"""
import argparse
import json
import sys

import pymupdf

COLORS = {
    "yellow":  (1.0, 1.0, 0.0),     # 时间类
    "orange":  (1.0, 0.75, 0.0),    # 特殊报名方式 / 原件要求
    "green":   (0.6, 0.93, 0.6),    # CA / 重点平台命中
    # 模块一扩展（v1.5）：投标资料递交阶段分类色
    "purple":  (0.78, 0.65, 0.95),  # 🟪 报名环节递交（紫）
    "blue":    (0.55, 0.78, 1.0),   # 🟦 投标文件递交（蓝）
    # v1.7：备查/中标后类——仅文字批注，不着色不高亮
    "none":    None,
}


def merge_rects(rects, gap=6.0):
    """将同一行上水平相邻（间距<gap）的矩形链式合并为一个，避免混排字体把
    一句话拆成几十个小矩形、生成几十个重叠高亮和弹窗批注。
    多行命中会按行各自合并（垂直中心距>=5 视为不同行）。"""
    merged = [pymupdf.Rect(r) for r in rects]
    changed = True
    while changed:
        changed = False
        merged.sort(key=lambda x: ((x.y0 + x.y1) / 2, x.x0))
        out = []
        for r in merged:
            for i, m in enumerate(out):
                if (abs((m.y0 + m.y1) / 2 - (r.y0 + r.y1) / 2) < 5
                        and (r.x0 - m.x1) < gap and (m.x0 - r.x1) < gap):
                    # 显式写回列表（Rect 不支持原地 |=）
                    out[i] = pymupdf.Rect(min(m.x0, r.x0), min(m.y0, r.y0),
                                          max(m.x1, r.x1), max(m.y1, r.y1))
                    changed = True
                    break
            else:
                out.append(r)
        merged = out
    return merged


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", required=True)
    args = ap.parse_args()

    with open(args.tasks, "r", encoding="utf-8") as f:
        spec = json.load(f)

    doc = pymupdf.open(spec["input"])
    report = {"output": spec["output"], "matched": [], "unmatched": []}

    for task in spec.get("tasks", []):
        search = task.get("search", "")
        if not search:
            continue
        color = COLORS.get(task.get("color", "yellow"), COLORS["yellow"])
        comment = task.get("comment", "")
        occurrence = task.get("occurrence")
        total = 0
        for pno in range(len(doc)):
            page = doc[pno]
            hits = merge_rects(page.search_for(search))
            for hi, rect in enumerate(hits):
                if occurrence is not None and (total + hi) != occurrence:
                    continue
                if color is not None:
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=color)
                    annot.set_info(
                        title="AI辅助分析(须人工复核)",
                        content=comment or search,
                    )
                    annot.update()
                    if comment:
                        # 弹窗式批注锚定在高亮处（每个合并后的高亮区仅一条）
                        tann = page.add_text_annot(rect.tl, comment)
                        tann.set_info(title="AI辅助分析(须人工复核)", content=comment)
                        tann.update()
                else:
                    # "none"：仅文字批注（备查/中标后类），不高亮
                    if comment:
                        tann = page.add_text_annot(rect.tl, comment)
                        tann.set_info(title="AI辅助分析(须人工复核)", content=comment)
                        tann.update()
            total += len(hits)
        if total > 0:
            report["matched"].append({"search": search, "hits": total,
                                      "color": task.get("color", "yellow")})
        else:
            report["unmatched"].append({"search": search,
                                        "color": task.get("color", "yellow"),
                                        "reason": "PDF原文中未找到该精确文本（可能跨行/跨页断开），需人工定位"})

    doc.save(spec["output"], garbage=3, deflate=True)
    doc.close()
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["unmatched"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
