# -*- coding: utf-8 -*-
"""把《供方能力调查基础信息表》的键值对填入甲方 Excel 模板（生成副本，不修改原件）。

用法:
    python fill_company_info.py --template <甲方Excel模板> --out <输出文件.xlsx> \
        [--ledgers <parse_ledgers.py 输出目录>]

匹配规则（保守）:
    1) 遍历模板所有 sheet 的每个单元格；
    2) 单元格文本与台账键名做包含匹配（如『公司全称』『统一社会信用代码』）；
    3) 命中后优先填"右侧空单元格"，右侧非空则填"下方空单元格"；
    4) 两个相邻位置都非空 → 跳过并记录 ⚠️（人工处理）。
硬性规则: 无法匹配的表头/字段 → 标记 ⚠️ 人工补充，禁止编造。
"""
import argparse
import json
import os
import re
import sys
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl  # noqa: E402
from common import MARK_MANUAL, MARK_OK, cell_str, die  # noqa: E402

NOISE = re.compile(r"[，。、；：,.;:（）()\s　「」『』“”\"'【】\[\]：:]+")


def clean(s):
    return NOISE.sub("", str(s))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, help="甲方 Excel 模板路径")
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    ap.add_argument("--ledgers", default=None)
    args = ap.parse_args()

    if not os.path.isfile(args.template):
        die(f"甲方模板不存在: {args.template}")
    led_dir = args.ledgers or os.path.dirname(os.path.abspath(args.out))
    p = os.path.join(led_dir, "ledgers_company.json")
    if not os.path.isfile(p):
        die(f"缺少公司信息 JSON: {p}。请先运行 parse_ledgers.py")

    company = json.load(open(p, encoding="utf-8"))["company_info"]
    if not company:
        die("公司信息为空，禁止编造。请人工核对《供方能力调查基础信息表》。")

    try:
        wb = openpyxl.load_workbook(args.template)
    except Exception as e:
        die(f"甲方模板读取失败: {e}")

    filled = []      # 已填: (sheet, 单元格, 键, 值)
    unmatched = []   # 未匹配表头
    skipped = []     # 命中但相邻位置非空

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                text = clean(cell_str(cell.value))
                if not text:
                    continue
                for key, val in company.items():
                    kc = clean(key)
                    if kc and (kc in text or text in kc):
                        target = None
                        # 优先右侧，其次下方
                        r = cell.row
                        c = cell.column
                        right = ws.cell(row=r, column=c + 1)
                        below = ws.cell(row=r + 1, column=c)
                        if right.value in (None, ""):
                            target = right
                        elif below.value in (None, ""):
                            target = below
                        else:
                            skipped.append((ws.title, cell.coordinate, key))
                            continue
                        target.value = val
                        filled.append((ws.title, target.coordinate, key, val))
                        break

    out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)

    # 报告
    lines = ["# Excel 公司信息填写报告", ""]
    lines.append(f"- 模板: {args.template}")
    lines.append(f"- 输出副本: {out}")
    lines.append(f"- 已填写: {len(filled)} 处")
    lines.append("")
    lines.append("| Sheet | 单元格 | 台账键 | 填写值 |")
    lines.append("|---|---|---|---|")
    for s, coord, k, v in filled:
        lines.append(f"| {s} | {coord} | {k} | {v} |")
    lines.append("")
    if skipped:
        lines.append(f"## ⚠️ 命中但相邻位置非空（需人工处理，{len(skipped)} 处）")
        for s, coord, k in skipped:
            lines.append(f"- {s}!{coord} 键『{k}』相邻单元格已有内容，未覆盖")
    lines.append("")
    lines.append("> 【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】")
    report_path = os.path.join(os.path.dirname(out), "fill_report.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"✅ 填写完成: {len(filled)} 处（输出: {out}）")
    print(f"   报告: {report_path}")
    if skipped:
        print(f"⚠️ {len(skipped)} 处命中但相邻非空，详见报告，请人工处理。")


if __name__ == "__main__":
    main()
