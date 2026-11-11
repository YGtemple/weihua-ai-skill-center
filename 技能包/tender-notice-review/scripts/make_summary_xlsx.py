# -*- coding: utf-8 -*-
"""
生成《招标公告关键信息汇总表.xlsx》（模块一输出）。
用法：
  python make_summary_xlsx.py --rows 行数据.json --out 输出路径.xlsx [--project 项目名]
行数据 JSON 格式（数组）：
[
  {"关键类型":"时间-投标文件递交截止时间","具体内容":"2025年10月15日 09:30",
   "原文位置":"第2页 第3段","匹配数据源":"公告原文","状态":"已标记黄色高亮+批注",
   "备注":"...","人工复核确认":"【待人工复核确认】"}
]
固定列：关键类型 | 具体内容 | 原文位置（页码/段落） | 匹配数据源 | 状态标记 | 备注 | 人工复核确认
表首行强制置顶免责声明。
"""
import argparse
import datetime as dt
import json

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HEADERS = ["关键类型", "具体内容", "原文位置（页码/段落）", "匹配数据源", "状态标记", "备注", "人工复核确认"]
DISCLAIMER = ("【强制声明】AI仅辅助分析，所有内容必须人工逐条复核原件后才可使用，"
              "禁止直接对外提交。未提及的信息已标注【公告未提及】，模糊信息标注【需人工复核原文】。")

THIN = Border(*[Side(style="thin", color="999999")] * 4)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--project", default="")
    args = ap.parse_args()

    with open(args.rows, "r", encoding="utf-8") as f:
        rows = json.load(f)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "关键信息汇总"

    # 免责声明置顶
    ws.cell(1, 1, DISCLAIMER)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    c = ws.cell(1, 1)
    c.font = Font(bold=True, color="9C0006", size=11)
    c.fill = PatternFill("solid", fgColor="FFC7CE")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 42

    # 元信息
    ws.cell(2, 1, f"项目名称：{args.project or '（未提供）'}    "
                  f"生成时间：{dt.datetime.now().strftime('%Y-%m-%d %H:%M')}    "
                  f"数据源：CA登记表.xlsx / 重点平台清单.xlsx（本地权威数据源，只读）")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEADERS))
    ws.cell(2, 1).font = Font(size=10, color="555555")

    # 表头
    hrow = 3
    for j, h in enumerate(HEADERS, start=1):
        c = ws.cell(hrow, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN

    # 数据行
    for i, r in enumerate(rows, start=hrow + 1):
        vals = [
            r.get("关键类型", ""),
            r.get("具体内容", ""),
            r.get("原文位置", ""),
            r.get("匹配数据源", ""),
            r.get("状态标记", ""),
            r.get("备注", ""),
            r.get("人工复核确认", "【待人工复核确认】"),
        ]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(i, j, v)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = THIN

    widths = [22, 40, 20, 22, 18, 30, 18]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = f"A{hrow + 1}"

    wb.save(args.out)
    print(f"已生成：{args.out}（共 {len(rows)} 条关键信息）")


if __name__ == "__main__":
    main()
