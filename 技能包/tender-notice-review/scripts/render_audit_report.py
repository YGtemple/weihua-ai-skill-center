# -*- coding: utf-8 -*-
"""
生成《报名资料合规审核报告.xlsx》（模块二输出）。
用法：
  python render_audit_report.py --data 审核数据.json --out 输出路径.xlsx
审核数据 JSON 格式：
{
  "project": "项目名称",
  "announcement": "公告文件名",
  "materials": "报名资料清单/文件夹路径",
  "six_dimensions": [
    {"维度":"漏项审核","结论":"...","明细":["..."]}
  ],
  "compliant":     [{"item":"营业执照","basis":"公告第X条","evidence":"..."}],
  "noncompliant":  [{"item":"缺少安全生产许可证","basis":"公告第X条要求...","suggestion":"补充..."}],
  "manual_review": [{"item":"资质专业范围匹配","reason":"语义近似无法精准匹配，需人工核对"}],
  "collected_files": [{"source":"原路径","dest":"归集路径"}]
}
报告结构：强制兜底声明置顶 → 总览 → ✅合规项 → ❌错误/缺失项 → ⚠️人工必复核项 → 六维度结论 → 归集文件清单。
"""
import argparse
import datetime as dt
import json

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DISCLAIMER = ("【强制兜底声明】本AI审核仅为辅助参考，存在识别盲区与匹配误差，所有资料、资质、有效期、"
              "报名要求务必人工对照公告原件及证书原件二次复核，确认无误后方可报名提交。")
THIN = Border(*[Side(style="thin", color="999999")] * 4)


def section_title(ws, row, text, color):
    ws.cell(row, 1, text)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = ws.cell(row, 1)
    c.font = Font(bold=True, size=12)
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(vertical="center")
    return row + 1


def table_header(ws, row, headers, color="305496"):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN
    return row + 1


def data_rows(ws, row, headers, items):
    for it in items:
        for j, h in enumerate(headers, start=1):
            v = it.get(h, "")
            c = ws.cell(row, j, v if v is not None else "")
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = THIN
        row += 1
    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        d = json.load(f)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "合规审核报告"

    # 1) 免责声明置顶
    ws.cell(1, 1, DISCLAIMER)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    c = ws.cell(1, 1)
    c.font = Font(bold=True, color="9C0006", size=11)
    c.fill = PatternFill("solid", fgColor="FFC7CE")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 48

    # 2) 基本信息
    info = [
        ("项目名称", d.get("project", "（未提供）")),
        ("审核依据-公告", d.get("announcement", "（未提供）")),
        ("审核依据-数据源", "CA登记表.xlsx / 重点平台清单.xlsx（本地权威数据源，只读）+ 公告原文"),
        ("我方报名资料", d.get("materials", "（未提供）")),
        ("审核时间", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("统计", f"✅合规 {len(d.get('compliant', []))} 项 | "
                f"❌错误/缺失 {len(d.get('noncompliant', []))} 项 | "
                f"⚠️人工必复核 {len(d.get('manual_review', []))} 项"),
    ]
    r = 3
    for k, v in info:
        ws.cell(r, 1, k).font = Font(bold=True)
        ws.cell(r, 2, v)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        r += 1
    r += 1

    # 3) ✅ 合规项
    r = section_title(ws, r, "一、✅ 合规项（客观比对通过，仍需人工抽检）", "C6EFCE")
    hdr = ["项目", "公告原文依据", "比对证据", "备注"]
    r = table_header(ws, r, hdr)
    r = data_rows(ws, r, hdr, d.get("compliant", []))
    r += 1

    # 4) ❌ 错误/缺失项
    r = section_title(ws, r, "二、❌ 错误/缺失项（附修改补充建议）", "FFC7CE")
    hdr = ["项目", "公告原文依据", "问题描述", "修改补充建议"]
    r = table_header(ws, r, hdr)
    r = data_rows(ws, r, hdr, d.get("noncompliant", []))
    r += 1

    # 5) ⚠️ 人工必复核项
    r = section_title(ws, r, "三、⚠️ 人工必复核项（无数据源或公告原文依据，不做判定）", "FFEB9C")
    hdr = ["项目", "需人工复核原因", "建议核对方式", "备注"]
    r = table_header(ws, r, hdr)
    r = data_rows(ws, r, hdr, d.get("manual_review", []))
    r += 1

    # 6) 六维度审核结论
    r = section_title(ws, r, "四、六大维度审核结论", "D9E1F2")
    hdr = ["审核维度", "结论", "明细"]
    r = table_header(ws, r, hdr)
    for item in d.get("six_dimensions", []):
        ws.cell(r, 1, item.get("维度", "")).border = THIN
        ws.cell(r, 2, item.get("结论", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).border = THIN
        ws.cell(r, 3, "\n".join(item.get("明细", []))).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 3).border = THIN
        r += 1
    r += 1

    # 7) 归集文件清单
    r = section_title(ws, r, "五、合规资料归集清单（仅复制，未移动/未修改原文件）", "E2EFDA")
    hdr = ["原文件路径", "归集后路径"]
    r = table_header(ws, r, hdr)
    r = data_rows(ws, r, hdr, d.get("collected_files", []))

    for j, w in enumerate([26, 40, 45, 40, 20, 18], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    wb.save(args.out)
    print(f"已生成：{args.out}")


if __name__ == "__main__":
    main()
