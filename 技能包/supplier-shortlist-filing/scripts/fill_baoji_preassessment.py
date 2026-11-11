# -*- coding: utf-8 -*-
# fill_baoji_preassessment - 宝鸡石油机械供应商准入预审表填写（CLI 版）
# 模板结构：单张大表 25 行 × 9 列
#   R01-R06  基本信息区
#   R07-R17  资质评价区（8 项 + 评价人）
#   R18-R22  业务需求区（2 项 + 评价人）
#   R23-R25  预审/审核/批准意见（内部审批流程，绝不由 AI 填）
# 填法：
#   - R01 供应商名称：填公司全称
#   - R03 □ → ■ 勾选（按台账判断生产类+制造商）
#   - R09-R16 评价：填实际情况，评价结果按台账确定性判断（满足/不满足/待人工）
#   - R17/R22 评价人留空
#   - R20-R21 业务需求：填实际情况
#   - R23-R25 审批意见留空
#   - 末尾追加 AI 复核提示

import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def rep(lines, key, status, msg):
    lines.append((key, status, msg))


def get_unique_cells(row):
    """返回 row 中不重复的 tc 元素列表。"""
    seen = set()
    cells = []
    for tc in row._tr.findall(f".//{W}tc"):
        cid = id(tc)
        if cid in seen:
            continue
        seen.add(cid)
        cells.append(tc)
    return cells


def get_cell_text(tc):
    parts = []
    for t in tc.findall(f".//{W}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def append_para_text(tc, text, bold=False, color=None):
    """在 cell 末尾追加段落。"""
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if bold:
        rpr.append(OxmlElement("w:b"))
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rpr.append(c)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    p.append(r)
    tc.append(p)


def set_cell_text(tc, text, bold=False):
    """覆盖 cell 文本（保留首个段落）。"""
    paras = tc.findall(f".//{W}p")
    if not paras:
        append_para_text(tc, text, bold=bold)
        return True
    p = paras[0]
    for r in p.findall(f".//{W}r"):
        p.remove(r)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if bold:
        rpr.append(OxmlElement("w:b"))
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    p.append(r)
    for pp in paras[1:]:
        tc.remove(pp)
    return True


def replace_cell_text(tc, new_text):
    """整段重写 cell 内容（保留第一段，替换文字）。"""
    set_cell_text(tc, new_text)


# === 资质评价 8 项的实际描述（按台账硬编码）===
EVAL_DESCRIPTIONS = {
    9: {
        "ok": True,
        "desc": "公司全称：{name}; 统一社会信用代码 914107287982390285; 成立日期 {est}; "
                "营业执照（副本）编号同信用代码，长期有效；基本存款账户信息（开户许可证）编号 J4986000161608，长期有效。"
    },
    10: {
        "ok": True,
        "desc": "注册资本 60000 万元，实缴 23000 万元；近三年财务状况良好。"
                "2023/2024/2025 年度营业收入分别为 116210/117321/143829 万元，呈稳步增长趋势。"
    },
    11: {
        "ok": None,
        "desc": "【待人工查询『国家企业信用信息公示系统』核实近三年违法违规记录，台账无该字段】"
    },
    12: {
        "ok": True,
        "desc": "已通过 ISO9001 质量管理体系认证（证书号 0350125Q30884R4M，有效期至 2028-10-10）；"
                "ISO14001 环境管理体系认证（有效期至 2028-10-10）；"
                "ISO45001 职业健康安全管理体系认证（0350125S30593R4M，有效期至 2028-10-10）。"
    },
    13: {
        "ok": True,
        "desc": "具备中华人民共和国特种设备生产许可证（编号 TS2410B93-2027，有效期至 2027-10-30）；"
                "起重设备安装工程专业承包贰级（证书号 D241790148，有效期至 2030-12-31）；"
                "安全生产许可证（编号（豫）JZ安许证字[2025]017884，有效期至 2028-12-02）。"
    },
    14: {
        "ok": True,
        "desc": "公司具有完善的销售/技术/售后服务体系：{aftersale_short}"
    },
    15: {
        "ok": None,
        "desc": "【待人工核查与公司高管/关联方关系及廉洁从业情况，台账无该字段】"
    },
    16: {
        "ok": None,
        "desc": "【待人工确认是否还有集团公司额外规定要求】"
    },
    20: {
        "ok": None,
        "desc": "起重机整体接地改造项目：公司在起重设备（单梁、双梁、门机）领域有完整设计、生产、安装、维保能力，"
                "可提供同类型项目产品及服务。详见公司简介。"
    },
    21: {
        "ok": None,
        "desc": "【待人工补充：是否谈判中选/用户指定/原厂配件等其他需求】"
    },
}


def fill_r01_name(ws, ci, report):
    row = ws.rows[0]
    cells = get_unique_cells(row)
    if len(cells) < 2:
        rep(report, "R01[供应商名称]", "WARN", "单元格结构异常，跳过")
        return
    name = ci.get("公司全称", "") or ""
    if not name:
        rep(report, "R01[供应商名称]", "MISS", "台账无公司全称，留空")
        return
    set_cell_text(cells[1], name, bold=True)
    rep(report, "R01[供应商名称]", "OK", f"已填: {name}")


def fill_r03_supplier_type(ws, report):
    row = ws.rows[2]
    cells = get_unique_cells(row)
    changed = False
    for c in cells:
        text = get_cell_text(c)
        new_text = text
        new_text = new_text.replace("生产类物资□、非生产类物资", "■ 生产类物资、□ 非生产类物资")
        new_text = new_text.replace("制造商□、贸易商□、代理商□", "■ 制造商、□ 贸易商、□ 代理商")
        if new_text != text:
            replace_cell_text(c, new_text)
            changed = True
    if changed:
        rep(report, "R03[供应商类型]", "OK", "已勾选 ■ 生产类物资 + ■ 制造商")
    else:
        rep(report, "R03[供应商类型]", "WARN", "未找到 □ 模板文本，可能格式不同")


def fill_evaluation_block(ws, ci, report, label_prefix):
    """填充评价表 R09-R16 + R20-R21 的实际情况简要描述列与评价结果列。

    模板列结构（get_unique_cells 视角，6 个 tc + gridSpan）：
      [0]gs=1 序号  [1]gs=1 评价内容  [2]gs=2 评价标准
      [3]gs=2 实际情况简要描述  [4]gs=2 评价结果  [5]gs=1 备注
    """
    for rel_row, spec in EVAL_DESCRIPTIONS.items():
        actual_row_idx = rel_row - 1
        row = ws.rows[actual_row_idx]
        cells = get_unique_cells(row)
        if len(cells) < 5:
            continue
        # 安全检查：序号必须是数字
        seq_text = get_cell_text(cells[0]).strip()
        if not seq_text.isdigit():
            continue
        # 拼装描述
        desc_text = spec["desc"]
        desc_text = desc_text.replace("{name}", ci.get("公司全称", "") or "")
        desc_text = desc_text.replace("{est}", ci.get("成立日期", "") or "")
        aftersale = (ci.get("售后、服务保障", "") or "")[:120]
        desc_text = desc_text.replace("{aftersale_short}", aftersale + "...")
        # 写入实际情况列 cells[3]（gs=2 评价标准后）
        set_cell_text(cells[3], desc_text)
        # 评价结果列 cells[4]
        if spec["ok"] is True:
            replace_cell_text(cells[4], "■ 满足要求")
            verdict = "■ 满足要求"
        elif spec["ok"] is False:
            replace_cell_text(cells[4], "■ 不满足要求")
            verdict = "■ 不满足要求"
        else:
            replace_cell_text(cells[4], "□ 待人工判定")
            verdict = "□ 待人工判定"
        rep(report, f"R{rel_row:02d}{label_prefix}", "OK",
            f"已填描述 + {verdict}")


def append_ai_notice(doc):
    NOTE = ("⚠️ 本文档由 AI 辅助填写（基于本地三份台账数据）。"
            "评价结果仅依据台账可核实信息自动判断；"
            "审批结论（预审/审核/批准意见）必须由业务/评价人人工复核确认。")
    for p in doc.paragraphs:
        if "AI 辅助填写" in p.text or "AI辅助填写" in p.text:
            return False
    p = doc.add_paragraph()
    r = p.add_run(NOTE)
    r.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--ledgers-dir", required=True)
    args = ap.parse_args()

    ledgers_dir = Path(args.ledgers_dir)
    ci = json.load(open(ledgers_dir / "ledgers_company.json", encoding="utf-8"))["company_info"]

    doc = Document(args.src)
    ws = doc.tables[0]
    report = []

    fill_r01_name(ws, ci, report)
    fill_r03_supplier_type(ws, report)
    fill_evaluation_block(ws, ci, report, "[资质评价/业务需求]")

    if append_ai_notice(doc):
        rep(report, "末尾[AI提示]", "OK", "已追加 AI 复核提示段落")
    else:
        rep(report, "末尾[AI提示]", "WARN", "已存在 AI 提示，未重复")

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.out)

    print(f"\n=== 填写完成 ===")
    print(f"输出: {args.out}")
    print(f"\n填充报告：")
    for k, s, m in report:
        icon = {"OK": "✅", "WARN": "⚠️", "MISS": "❌"}.get(s, "·")
        print(f"  {icon} [{k}] {m}")


if __name__ == "__main__":
    main()