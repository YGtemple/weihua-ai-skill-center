# -*- coding: utf-8 -*-
"""
生成《报名资料合规审核报告.xlsx》(v1.5 升级)。

强制结构：
- 主 sheet「合规审核报告」：强制兜底声明置顶 + 项目基本信息 + 5 大板块
- 独立 sheet「人工资料完整性核对对照表」(v1.4 升级 + v1.5 升级)：
  - 基础分块：PDF 后附扫描线索
  - 升级分块：截图类资料页面锚点识别线索（v1.5 新增）

v1.5 升级项：
- 「⚠️ 人工必复核项」由六分类升级为七分类，新增"📷 截图类资料人工必复核"；
- 截图类证明资料（营业执照、中国执行信息公开网、信用中国截图等）严格按
  references/screenshot-proof-rules.md 处理：禁止判定 ❌ 漏项，统一归入 ⚠️ 人工必复核。

用法：
  python render_audit_report.py --data 审核数据.json --out 输出路径.xlsx

审核数据 JSON 格式参见 references/audit-output-schema.md：
{
  "project": "项目名称",
  "announcement": "公告文件名",
  "summary_table": "来自汇总表/直接解析原始招标公告",
  "mode": "A" 或 "B",
  "materials_bidding": "报名资料文件夹/文件列表",
  "materials_tendering": "投标资料文件夹/文件列表",
  "compliant":        [{"item":"...","basis":"...","evidence":"...","note":"..."}],
  "missing_bidding":  [{"item":"...","basis":"...","problem":"...","suggestion":"..."}],
  "wrong_bidding":    [{"item":"...","basis":"...","problem":"...","suggestion":"..."}],
  "missing_tendering":[{"item":"...","basis":"...","problem":"...","suggestion":"..."}],
  "wrong_tendering":  [{"item":"...","basis":"...","problem":"...","suggestion":"..."}],
  "manual_review":    [{"category":"🟩现场备查","item":"...","basis":"...","advice":"..."}],
  "six_dimensions":   [{"dimension":"...","conclusion":"...","details":[...]}],
  "collected_files":  [{"stage":"报名","source":"...","dest":"..."}],
  "houfu_scan":       { 来自 scan_pdf_houfu.py 的输出 }
}
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DISCLAIMER = ("【强制兜底声明】本AI审核仅为辅助参考，存在识别盲区与匹配误差，"
              "务必区分【报名递交资料】【投标递交资料】【现场备查】【中标后提交】，"
              "所有资料、资质、有效期、报名要求务必人工对照公告原件及证书原件二次复核，"
              "确认无误后方可提交。")

THIN = Border(*[Side(style="thin", color="999999")] * 4)
GREEN_HEADER_FILL = PatternFill("solid", fgColor="70AD47")
RED_HEADER_FILL = PatternFill("solid", fgColor="C00000")
YELLOW_HEADER_FILL = PatternFill("solid", fgColor="FFC000")
BLUE_HEADER_FILL = PatternFill("solid", fgColor="305496")
PURPLE_HEADER_FILL = PatternFill("solid", fgColor="7030A0")


def section_title(ws, row, text, color, span=6):
    ws.cell(row, 1, text)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row, 1)
    c.font = Font(bold=True, size=12, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    return row + 1


def table_header(ws, row, headers, color=BLUE_HEADER_FILL):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = color
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


def render_main_sheet(ws, d):
    # 1) 免责声明置顶
    ws.cell(1, 1, DISCLAIMER)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    c = ws.cell(1, 1)
    c.font = Font(bold=True, color="9C0006", size=11)
    c.fill = PatternFill("solid", fgColor="FFC7CE")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 60

    # 2) 项目基本信息
    mode_label = "模式 A：来自【招标公告智能标记解析】汇总表" if d.get("mode") == "A" else (
        "模式 B：直接解析原始招标公告【存在识别偏差风险】" if d.get("mode") == "B" else "未指定"
    )
    counts = {
        "✅": len(d.get("compliant", [])),
        "报名漏项": len(d.get("missing_bidding", [])),
        "报名错误": len(d.get("wrong_bidding", [])),
        "投标漏项": len(d.get("missing_tendering", [])),
        "投标错误": len(d.get("wrong_tendering", [])),
        "人工必复核": len(d.get("manual_review", [])),
    }

    info = [
        ("项目名称", d.get("project", "（未提供）")),
        ("审核依据-公告", d.get("announcement", "（未提供）")),
        ("审核依据-汇总表", d.get("summary_table", "（未提供）")),
        ("模式", mode_label),
        ("我方报名资料", d.get("materials_bidding", "（未提供）")),
        ("我方投标资料", d.get("materials_tendering", "（未提供）")),
        ("数据源", "CA登记表.xlsx / 重点平台清单.xlsx（本地权威数据源，只读）+ 公告原文"),
        ("审核时间", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("统计",
         f"✅合规 {counts['✅']} 项 | "
         f"❌报名漏项 {counts['报名漏项']} / 报名错误 {counts['报名错误']} / "
         f"投标漏项 {counts['投标漏项']} / 投标错误 {counts['投标错误']} | "
         f"⚠️人工必复核 {counts['人工必复核']} 项"),
    ]
    r = 3
    for k, v in info:
        ws.cell(r, 1, k).font = Font(bold=True)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor="F2F2F2")
        ws.cell(r, 1).alignment = Alignment(vertical="top")
        ws.cell(r, 1).border = THIN
        ws.cell(r, 2, v).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.cell(r, 2).border = THIN
        # 当内容很长时给个估的行高
        if isinstance(v, str) and len(v) > 80:
            ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 18, 32)
        r += 1
    r += 1

    # 3) ✅ 合规项
    r = section_title(ws, r, "一、✅ 合规项（客观比对通过，仍需人工抽检）", "70AD47")
    hdr = ["项目", "公告原文依据", "比对证据", "备注"]
    r = table_header(ws, r, hdr, GREEN_HEADER_FILL)
    r = data_rows(ws, r, hdr, d.get("compliant", []))
    r += 1

    # 4) ❌ 错误/缺失项（强制四分类）
    r = section_title(ws, r, "二、❌ 错误/缺失项（按四阶段分类，附修改补充建议）", "C00000")
    r = _render_noncompliant_subtable(ws, r, "❌【报名漏项】", d.get("missing_bidding", []))
    r = _render_noncompliant_subtable(ws, r, "❌【报名错误】", d.get("wrong_bidding", []))
    r = _render_noncompliant_subtable(ws, r, "❌【投标漏项】", d.get("missing_tendering", []))
    r = _render_noncompliant_subtable(ws, r, "❌【投标错误】", d.get("wrong_tendering", []))
    r += 1

    # 5) ⚠️ 人工必复核项
    r = section_title(ws, r, "三、⚠️ 人工必复核项（无数据源/无原文依据/阶段特殊/后附扫描线索）", "FFC000")
    hdr = ["类别", "条目", "原文要求", "建议核对方式"]
    r = table_header(ws, r, hdr, YELLOW_HEADER_FILL)
    r = data_rows(ws, r, hdr, _maybe_wrap_category(d.get("manual_review", [])))
    r += 1

    # 6) 六维度结论
    r = section_title(ws, r, "四、六大维度审核结论", "305496")
    hdr = ["审核维度", "结论", "明细"]
    r = table_header(ws, r, hdr, BLUE_HEADER_FILL)
    for item in d.get("six_dimensions", []):
        ws.cell(r, 1, item.get("dimension", "")).border = THIN
        ws.cell(r, 1).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(r, 2, item.get("conclusion", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).border = THIN
        ws.cell(r, 3, "\n".join(item.get("details", []))).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 3).border = THIN
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws.cell(r, 3).border = THIN
        r += 1
    r += 1

    # 7) 归集文件清单
    r = section_title(ws, r, "五、合规资料归集清单（仅复制，未移动/未修改原文件）", "7030A0")
    hdr = ["阶段", "原文件路径", "归集后路径"]
    r = table_header(ws, r, hdr, PURPLE_HEADER_FILL)
    r = data_rows(ws, r, hdr, d.get("collected_files", []))

    # 列宽
    widths = [22, 38, 45, 38, 22, 30]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w


def _render_noncompliant_subtable(ws, row, subtitle, items):
    """❌ 子标签块：先写子标题，再写表头，再写数据。"""
    ws.cell(row, 1, subtitle)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = ws.cell(row, 1)
    c.font = Font(bold=True, color="FFFFFF", size=11)
    c.fill = PatternFill("solid", fgColor="E97132")
    c.alignment = Alignment(horizontal="left", vertical="center")
    row += 1
    hdr = ["项目", "公告原文依据", "问题描述", "修改/补充建议"]
    for j, h in enumerate(hdr, start=1):
        cell = ws.cell(row, j, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C65911")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
    row += 1
    # 数据行
    formatted = [{"项目": it.get("item", ""),
                  "公告原文依据": it.get("basis", ""),
                  "问题描述": it.get("problem", ""),
                  "修改/补充建议": it.get("suggestion", "")}
                 for it in items]
    return data_rows(ws, row, hdr, formatted)


def _maybe_wrap_category(items):
    """把 {'category': '🟩现场备查', ...} 兼容到老字段（直接通过）；这里做归一化。"""
    out = []
    for it in items:
        out.append({
            "类别": it.get("category") or it.get("类别") or "⚠️",
            "条目": it.get("item") or it.get("条目") or "",
            "原文要求": it.get("basis") or it.get("原文要求") or "",
            "建议核对方式": it.get("advice") or it.get("建议核对方式") or "",
        })
    return out


def render_houfu_sheet(ws, houfu):
    """独立 sheet「人工资料完整性核对对照表」(v1.4 强制)。

    数据来自 houfu_scan.json 的 results[].hits[] 展开。
    """
    DISCLAIMER_HOUFU = ("【强制声明】本sheet中所有条目均为 AI 扫描线索与告警提示，"
                        "AI 不对任何条目判定 ❌ 缺失，禁止直接对外提交；"
                        "所有条目必须人工打开原 PDF 复核后方可认定。")

    ws.cell(1, 1, DISCLAIMER_HOUFU)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    c = ws.cell(1, 1)
    c.font = Font(bold=True, color="9C0006", size=11)
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 48

    # 基本信息
    if houfu:
        info = [
            ("扫描时间", houfu.get("generated_at", "")),
            ("扫描窗口", f"向后 {houfu.get('window_pages', 5)} 页"),
            ("扫描 PDF 数", houfu.get("stats", {}).get("pdfs_scanned", 0) if isinstance(houfu.get("stats"), dict) else 0),
            ("命中关键词总次数", houfu.get("stats", {}).get("total_hits", 0) if isinstance(houfu.get("stats"), dict) else 0),
            ("✅ 检索到附件线索", houfu.get("stats", {}).get("found", 0) if isinstance(houfu.get("stats"), dict) else 0),
            ("⚠️ 风险提醒", houfu.get("stats", {}).get("risk", 0) if isinstance(houfu.get("stats"), dict) else 0),
            ("📷 扫描PDF无法识别", houfu.get("stats", {}).get("scan_unrecognized", 0) if isinstance(houfu.get("stats"), dict) else 0),
        ]
    else:
        info = [("本sheet说明", "未提供 houfu_scan.json，无后附扫描数据；本 sheet 仅作占位。")]

    r = 3
    for k, v in info:
        ws.cell(r, 1, k).font = Font(bold=True)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor="F2F2F2")
        ws.cell(r, 1).border = THIN
        ws.cell(r, 1).alignment = Alignment(vertical="top")
        ws.cell(r, 2, v).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).border = THIN
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        r += 1
    r += 1

    # 表头
    headers = ["序号", "PDF 文件名", "PDF 总页数", "命中关键词", "后附标注所在页",
               "5 页窗口扫描范围", "结果分类", "命中附件线索", "备注", "建议"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(r, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN
    r += 1

    if not houfu:
        # 占位提示
        ws.cell(r, 1, "—").alignment = Alignment(horizontal="center")
        ws.cell(r, 2, "（无后附扫描数据）").alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        for j in range(1, 11):
            ws.cell(r, j).border = THIN
    else:
        idx = 1
        for pdf_r in houfu.get("results", []):
            pdf_name = pdf_r.get("pdf_name", "") or pdf_r.get("pdf", "")
            total_pages = pdf_r.get("total_pages", 0)
            hits = pdf_r.get("hits", [])
            if not hits:
                # 即便一份 PDF 没命中，也保留一行占位，让人工能立刻看到这份 PDF 已扫描
                ws.cell(r, 1, idx).alignment = Alignment(horizontal="center")
                ws.cell(r, 1).border = THIN
                ws.cell(r, 2, pdf_name).alignment = Alignment(wrap_text=True, vertical="top")
                ws.cell(r, 2).border = THIN
                ws.cell(r, 3, total_pages).alignment = Alignment(horizontal="center")
                ws.cell(r, 3).border = THIN
                ws.cell(r, 4, "（本PDF未命中'XXX：后附'关键词，跳过扫描）").alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=10)
                for j in [4]:
                    ws.cell(r, j).border = THIN
                r += 1
                idx += 1
                continue

            for h in hits:
                ws.cell(r, 1, idx).alignment = Alignment(horizontal="center")
                ws.cell(r, 1).border = THIN
                ws.cell(r, 2, pdf_name).alignment = Alignment(wrap_text=True, vertical="top")
                ws.cell(r, 2).border = THIN
                ws.cell(r, 3, total_pages).alignment = Alignment(horizontal="center")
                ws.cell(r, 3).border = THIN
                ws.cell(r, 4, h.get("keyword", "")).alignment = Alignment(wrap_text=True, vertical="top")
                ws.cell(r, 4).border = THIN
                ws.cell(r, 5, h.get("page", "")).alignment = Alignment(horizontal="center")
                ws.cell(r, 5).border = THIN
                ws.cell(r, 6, h.get("window", "")).alignment = Alignment(horizontal="center", wrap_text=True)
                ws.cell(r, 6).border = THIN
                # 结果分类上色
                cat_cell = ws.cell(r, 7, h.get("result", ""))
                cat_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cat_cell.border = THIN
                if h.get("result", "").startswith("✅"):
                    cat_cell.fill = PatternFill("solid", fgColor="C6EFCE")
                elif h.get("result", "").startswith("⚠️"):
                    cat_cell.fill = PatternFill("solid", fgColor="FFC7CE")
                elif h.get("result", "").startswith("📷"):
                    cat_cell.fill = PatternFill("solid", fgColor="FFEB9C")
                ws.cell(r, 8, h.get("evidence", "")).alignment = Alignment(wrap_text=True, vertical="top")
                ws.cell(r, 8).border = THIN
                ws.cell(r, 9, h.get("note", "")).alignment = Alignment(wrap_text=True, vertical="top")
                ws.cell(r, 9).border = THIN
                ws.cell(r, 10, h.get("advice", "")).alignment = Alignment(wrap_text=True, vertical="top")
                ws.cell(r, 10).border = THIN
                r += 1
                idx += 1

    # 列宽
    widths = [6, 28, 10, 28, 10, 22, 22, 32, 50, 36]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w


def render_screenshot_proofs_sheet(ws, screenshot_proofs):
    """v1.5 新增：在「人工资料完整性核对对照表」sheet 末尾追加截图类资料页面锚点识别线索分块。

    screenshot_proofs JSON 格式：
    [
      {
        "item": "营业执照",
        "pdf": "报名资料.pdf",
        "anchor_status": "✅ 已检测到资料页面（图片/截图）" | "⚠️ 未检索到锚点关键词" | "📷 PDF 文本提取为空（疑似整本扫描件）",
        "anchor_keyword": "统一社会信用代码",     # 仅路径 A 必填
        "anchor_page": 3,                         # 仅路径 A 必填
        "note": "已检测到对应资料页面；页面内容为图片/截图，无法读取图片内部文字，请人工复核...",
        "advice": "人��打开 PDF 第 3 页复核图片内容"
      }
    ]
    """
    if not screenshot_proofs:
        # 无数据时仅写一行占位说明
        r = ws.max_row + 3
        ws.cell(r, 1, "【v1.5 新增】截图类资料页面锚点识别线索（无数据）").font = Font(bold=True, color="305496", size=11)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        ws.row_dimensions[r].height = 22
        r += 1
        ws.cell(r, 1, "—").alignment = Alignment(horizontal="center")
        ws.cell(r, 2, "（本批次无截图类资料页面锚点识别线索；如 PDF 整本为扫描件，请人工逐页翻阅确认）")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        for j in range(1, 11):
            ws.cell(r, j).border = THIN
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
        return

    r = ws.max_row + 3
    # 分块标题
    ws.cell(r, 1, "【v1.5 新增】截图类资料页面锚点识别线索")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    c = ws.cell(r, 1)
    c.font = Font(bold=True, color="FFFFFF", size=12)
    c.fill = PatternFill("solid", fgColor="305496")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 22
    r += 1

    # 强制声明
    ws.cell(r, 1, ("【强制声明】截图类资料页面内容为图片/截图，AI 文本层无法读取图片内部文字；"
                   "本sheet仅记录 PDF 文本层是否检索到锚点关键词，不做 ❌ 缺失判定；"
                   "请人工打开PDF对应页面复核图片主体、日期、有效期。"))
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    c = ws.cell(r, 1)
    c.font = Font(bold=True, color="9C0006", size=10)
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 48
    r += 1

    # 表头
    headers = ["序号", "资料类型", "PDF 文件名", "锚点命中情况", "命中关键词", "命中页码", "备注", "建议核对方式"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(r, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN
    r += 1

    # 数据行
    for idx, p in enumerate(screenshot_proofs, start=1):
        ws.cell(r, 1, idx).alignment = Alignment(horizontal="center")
        ws.cell(r, 1).border = THIN
        ws.cell(r, 2, p.get("item", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).border = THIN
        ws.cell(r, 3, p.get("pdf", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 3).border = THIN
        # 锚点命中情况上色
        cat_cell = ws.cell(r, 4, p.get("anchor_status", ""))
        cat_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cat_cell.border = THIN
        if p.get("anchor_status", "").startswith("✅"):
            cat_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        elif p.get("anchor_status", "").startswith("⚠️"):
            cat_cell.fill = PatternFill("solid", fgColor="FFC7CE")
        elif p.get("anchor_status", "").startswith("📷"):
            cat_cell.fill = PatternFill("solid", fgColor="FFEB9C")
        ws.cell(r, 5, p.get("anchor_keyword", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 5).border = THIN
        ws.cell(r, 6, p.get("anchor_page", "")).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(r, 6).border = THIN
        ws.cell(r, 7, p.get("note", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 7).border = THIN
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        ws.cell(r, 7).border = THIN
        ws.cell(r, 9, p.get("advice", "")).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 9).border = THIN
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        ws.cell(r, 9).border = THIN
        r += 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        d = json.load(f)

    wb = openpyxl.Workbook()
    ws_main = wb.active
    ws_main.title = "合规审核报告"
    render_main_sheet(ws_main, d)

    # v1.4 + v1.5: 独立 sheet「人工资料完整性核对对照表」（基础分块 + 截图类资料升级分块）
    ws_houfu = wb.create_sheet("人工资料完整性核对对照表")
    render_houfu_sheet(ws_houfu, d.get("houfu_scan"))
    render_screenshot_proofs_sheet(ws_houfu, d.get("screenshot_proofs"))

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    print(f"��生成：{out_path}")


if __name__ == "__main__":
    main()
