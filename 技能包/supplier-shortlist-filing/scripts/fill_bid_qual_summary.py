# -*- coding: utf-8 -*-
"""投标单位资格预审汇总表 专用填充脚本（模块1：供应商入围资料库）

适用: 甲方《投标单位资格预审汇总表.xlsx》——单表、单数据行、横向 27 列宽表，
     部分列(序号/名称/注册/资本/法人/规模)已预填，需补齐企业简介/官网/邮箱/地址/
     厂房面积/资质证书等列。

用法:
    python fill_bid_qual_summary.py --template <甲方.xlsx> --out <成品.xlsx> --ledgers <台账JSON目录>

列映射为 2026-09-04 实测（新模板需按表头微调）:
    H  企业简介        <- 台账[简介]
    I  官网            <- 台账[网址]
    J  邮箱            <- 台账[邮箱号]
    K  地址            <- 台账[实际经营地址](与注册地址一致)
    S  厂房面积(㎡)     <- 台账[生产/服务场地面积] 85万平方米 → 850000
    Z  资质证书        <- 《资质清单.xlsx》证书完整名称(排除银行开户/税务) 每行一条

台账无数据/值异常 → 留空不编造，输出 ⚠️ 待人工清单。
"""
import argparse, json, os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--ledgers", required=True, help="台账JSON目录(parse_ledgers.py输出)")
    args = ap.parse_args()

    ci = json.load(open(os.path.join(args.ledgers, "ledgers_company.json"), encoding="utf-8"))["company_info"]
    qs = json.load(open(os.path.join(args.ledgers, "ledgers_qualifications.json"), encoding="utf-8"))
    if isinstance(qs, dict):
        qs = qs.get("qualifications", [])

    def g(k):
        return ci.get(k, "") or ""

    wb = load_workbook(args.template)  # 保留样式
    ws = wb.worksheets[0]
    fill = []

    def set_cell(addr, val, note):
        ws[addr] = val
        ws[addr].alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        fill.append((addr, "✅", note))

    # ===== 可权威填充列 =====
    set_cell("H3", g("简介"), "来源:简介")
    set_cell("I3", g("网址"), "来源:网址")
    set_cell("J3", g("邮箱号"), "来源:邮箱号")
    set_cell("K3", g("实际经营地址"), "来源:实际经营地址(与注册地址一致)")
    set_cell("S3", 850000, f"来源:生产/服务场地面积 {g('生产/服务场地面积')} → 850000㎡")

    skip = {"银行开户许可", "税务资质"}
    cert_names = [q.get("证书完整名称", "") for q in qs
                  if q.get("类别") not in skip and q.get("证书完整名称")]
    set_cell("Z3", "\n".join(cert_names), f"来源:《资质清单.xlsx》共{len(cert_names)}项")

    # 行高容纳多行(简介/资质)
    ws.row_dimensions[3].height = 150
    wb.save(args.out)

    print("已保存:", args.out)
    print("===== 填充报告 =====")
    for addr, st, note in fill:
        print(f"[{st}] {addr}: {note}")

    print("===== 留空待人工(台账无数据或值异常, 禁止编造) =====")
    for addr, name, reason in [
        ("L3", "营业收入/2025年(元)", "台账无财务数值键, 财务报表PDF禁止读取内容"),
        ("M3", "营业收入/2024年(元)", "同上"),
        ("N3", "营业收入/2023年(元)", "同上"),
        ("O3", "参保人数", "台账『缴纳社保人员』值异常(1902-01-26 疑似日期误录)"),
        ("P3", "员工总数", "台账『公司总人数』值异常(1907-02-12); 简介口径‘员工2600多人’仅供参考不写入"),
        ("Q3", "高级职称", "台账无职称人数数据"),
        ("R3", "中级职称", "台账无职称人数数据"),
        ("T3", "净利润率25年", "需财务数据, 台账无"),
        ("U3", "资产负债率25年", "需财务数据, 台账无"),
        ("V3", "人均营业收入(万元)", "依赖营业收入, 缺"),
        ("W3", "近三年同类项目业绩业主", "原占位符保留; 明细台账按约定未读取"),
        ("X3", "业绩金额合计(万元)", "同上, 缺"),
        ("Y3", "是否缴纳保证金", "台账无此信息"),
    ]:
        print(f"[⚠️] {addr} {name}: {reason}")

if __name__ == "__main__":
    main()
