# -*- coding: utf-8 -*-
"""供应商能力调查表填充 + 资质放置（场景 A2 基础填充）

用法:
    python fill_survey_template.py --src <甲方.docx> --out <成品.docx> --ledgers <台账JSON目录>
    # 不传参时回退到旧默认路径（向后兼容）

适用模板结构（2026-09 实测）:
    Table1 = 公司基本情况 24 行「项目-内容」两列表（部分行跨4列）
    Table2 = 公司认证情况 6 行（评审内容 + 认证机构/证书编号/有效期 标签）
    正文标题段: 企业简介 / 八、生产过程质量控制描述 / 售后、服务保障
    末尾资质标题段: 营业执照 / 特种设备生产许可证 / 质量管理体系证书 ...
    Table3/Table4(客户/设备明细) 由 fill_multi_sheet_tables.py（场景A3）另行处理

映射规则（2026-09 修订）:
    - 表1 行→台账键见 map1；联系电话/邮政编码 直接取台账键（不再当缺失）
    - 年度销售总额 → 台账键「上年度销售额」（源有值则填，禁止用财务 sheet 营业成本替代）
    - 员工情况 → 员工人数/管理人员数量/质检人员数量/本科以上学历 + 口径待核对备注
    - 表2 认证 6 行 → 发证单位/证书编号/有效期 三键追加到标签后
"""
import argparse
import json
import os
import sys

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.text.paragraph import Paragraph

NOTE = "【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】"

DEFAULT_WORK = r"C:\\卫华投标工作区\\_work"


def new_para_after(anchor_el):
    """在 anchor_el 之后插入一个空段落，返回段落 xml。"""
    p_xml = parse_xml('<w:p %s/>' % nsdecls('w'))
    anchor_el.addnext(p_xml)
    return p_xml


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=os.path.join(DEFAULT_WORK, "供方能力调查表模板(2)1.docx"))
    ap.add_argument("--out", default=os.path.join(DEFAULT_WORK, "供方能力调查表-已填写.docx"))
    ap.add_argument("--ledgers", default=DEFAULT_WORK)
    args = ap.parse_args()

    ci = json.load(open(os.path.join(args.ledgers, "ledgers_company.json"), encoding="utf-8"))["company_info"]
    quals = json.load(open(os.path.join(args.ledgers, "ledgers_qualifications.json"), encoding="utf-8"))["qualifications"]

    def g(k):
        return ci.get(k, "")

    fill_report = []

    def rep(pos, st, note):
        fill_report.append((pos, st, note))

    # ========== 1) 表1 公司基本情况 24 行映射（行→台账键；None 表示台账缺/异常需人工） ==========
    map1 = [
        (0, "公司全称", None),
        (1, "成立日期", None),
        (2, "企业性质", None),
        (3, "注册资本", None),
        (4, "法定代表人", None),
        (5, "联系电话", None),
        (6, "统一社会信用代码", None),
        (7, "公司地址", None),
        (8, "邮政编码", None),
        (9, "经营范围", None),
        (10, "产品", None),
        (11, "制造标准", None),
        (12, "供货能力", None),
        (13, "供货周期", None),
        (14, "主要产品质量状况描述", None),
        (15, "生产能力", None),
        (16, None, "年度销售总额取台账键:上年度销售额"),  # 特殊处理
        (17, None, "组合资质清单文本"),                      # 特殊处理
        (18, "公司面积", None),
        (19, None, "员工情况组合文本"),                      # 特殊处理
        (20, "本科以上学历", None),
        (21, "开户银行", None),
        (22, "账号", None),
        (23, None, "开户地址台账缺失,需人工补充"),
    ]

    def is_bad_date(v):
        return isinstance(v, str) and len(v) == 10 and "-" in v and v[:4].isdigit() and int(v[:4]) < 1980

    qual_summary = "；".join(
        f"{q.get('证书完整名称', '')}（{q.get('证书编号', '')}）"
        for q in quals if q.get('类别') not in ('银行开户许可', '税务资质'))

    # 员工情况：只放口径明确的字符串键；数字类键（38/757）口径待核对 → 备注
    staff_note_parts = []
    if g("技术人员数量"):
        staff_note_parts.append(f"技术人员数量{g('技术人员数量')}人")
    if g("缴纳社保人员"):
        staff_note_parts.append(f"缴纳社保人员{g('缴纳社保人员')}人")
    staff_txt = (f"员工总数{g('员工人数')}；管理人员{g('管理人员数量')}；质检人员{g('质检人员数量')}；"
                 f"本科以上学历{g('本科以上学历')}"
                 + (f"（{'、'.join(staff_note_parts)}台账口径待人工核对确认是否计入）" if staff_note_parts else ""))

    doc = Document(args.src)
    t1 = doc.tables[0]

    for ri, key, tip in map1:
        cell = t1.rows[ri].cells[1]
        row_label = t1.rows[ri].cells[0].text
        if key is None and tip and "资质清单" in tip:
            val, note, st = qual_summary, "来自《资质清单.xlsx》(排除银行开户/税务)", "OK"
        elif key is None and tip and "员工情况" in tip:
            val, note, st = staff_txt, "来自《供方能力调查基础信息表》", "OK"
        elif key is None and tip and "年度销售总额" in tip:
            v = g("上年度销售额")
            if v:
                val, note, st = v, "台账键:上年度销售额", "OK"
                # 与财务 sheet 口径差异提示（读不到 extended 就跳过）
                try:
                    ext = json.load(open(os.path.join(args.ledgers, "ledgers_company_extended.json"), encoding="utf-8"))
                    fin = ext["sheets"]["财务信息"]["metrics"]
                    rev = next((m for m in fin if m["name"].strip() == "营业收入"), None)
                    if rev:
                        years = sorted(rev["values"])
                        last_v = rev["values"].get(years[-1]) if years else None
                        if last_v:
                            note += f" | 注:财务sheet『营业收入』{years[-1]}={last_v}万元, 口径不同请人工复核"
                except Exception:
                    pass
            else:
                val, note, st = "", tip, "⚠️"
        elif key:
            v = g(key)
            if v and not is_bad_date(v):
                val, note, st = v, f"台账键:{key}", "OK"
            else:
                val, note, st = "", tip or f"台账键:{key}无有效值", "⚠️"
        else:
            val, note, st = "", tip or "", "⚠️"
        if val:
            p = cell.paragraphs[0]
            for r in list(p.runs):
                r.text = ""
            p.add_run(val)
        rep(f"表1-第{ri+1}行[{row_label}]", st,
            (val[:60] + ("…" if len(val) > 60 else "")) if val else (note or ""))

    # ========== 2) 表2 认证情况 6 项 ==========
    t2 = doc.tables[1]
    cert_map = [
        (1, "质量管理体系证书发证单位", "质量管理体系证书编号", "质量管理体系证书有效期", "认证机构"),
        (2, "环境管理体系证书发证单位", "环境管理体系证书编号", "环境管理体系证书有效期", "认证机构"),
        (3, "职业健康安全管理体系证书发证单位", "职业健康安全管理体系证书编号", "职业健康安全管理体系证书有效期", "认证机构"),
        (4, "建筑业企业资质证书证书发证单位", "建筑业企业资质证书证书编号", "建筑业企业资质证书证书有效期", "发证机关"),
        (5, "中华人民共和国特种设备生产许可证证书发证单位", "中华人民共和国特种设备生产许可证证书编号", "中华人民共和国特种设备生产许可证证书有效期", "发证机构"),
        (6, "安全生产许可证证书发证单位", "安全生产许可证证书编号", "安全生产许可证证书有效期", "发证机构"),
    ]
    for ri, k_org, k_no, k_valid, label in cert_map:
        cell = t2.rows[ri].cells[2]
        org, no, valid = g(k_org), g(k_no), g(k_valid)
        for p in cell.paragraphs:
            txt = p.text.strip()
            if txt.startswith("认证机构") or txt.startswith("发证机构") or txt.startswith("发证机关"):
                if org:
                    p.add_run(" " + org)
            elif txt.startswith("证书编号"):
                if no:
                    p.add_run(" " + no)
            elif txt.startswith("有 效 期") or txt.startswith("有效期至"):
                if valid:
                    p.add_run(" " + valid)
        status = "OK" if (org and no and valid) else "⚠️"
        rep(f"表2-第{ri+1}行[{t2.rows[ri].cells[0].text}]", status,
            f"机构:{org or '缺'} 编号:{no or '缺'} 有效期:{valid or '缺'}")

    # ========== 3) 简介/质控/售后段落 ==========
    def fill_after_title(title, content):
        if not content:
            rep(f"段落[{title}]", "⚠️", "台账无内容")
            return
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip() == title:
                nxt = doc.paragraphs[i + 1] if i + 1 < len(doc.paragraphs) else None
                if nxt and not nxt.text.strip():
                    for r in list(nxt.runs):
                        r.text = ""
                    nxt.add_run(content)
                    rep(f"段落[{title}]下方", "OK", "已填内容")
                else:
                    p_xml = new_para_after(p._p)
                    para = Paragraph(p_xml, doc)
                    para.add_run(content)
                    rep(f"段落[{title}]下方(新增)", "OK", "已填内容")
                return
        rep(f"段落[{title}]", "⚠️", "未找到标题段")

    fill_after_title("企业简介", g("简介"))
    fill_after_title("八、生产过程质量控制描述", g("质量控制描述"))
    fill_after_title("售后、服务保障", g("售后、服务保障"))

    # ========== 4) 资质放置到对应标题下 ==========
    def find_qual(cat):
        for q in quals:
            if q.get("类别") == cat:
                return q
        return None

    def insert_under_title(title, q):
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip() == title:
                img_path = q.get("_resolved_path", "") if q else ""
                if not img_path or not os.path.exists(img_path):
                    rep(f"资质[{title}]", "❌", f"原件不存在: {img_path or '台账未记录'}")
                    return False
                new_p = new_para_after(p._p)
                para = Paragraph(new_p, doc)
                run = para.add_run()
                try:
                    run.add_picture(img_path, width=Inches(5.0))
                except Exception as e:
                    rep(f"资质[{title}]", "❌", f"插图失败: {e}")
                    return False
                rep(f"资质[{title}]", "✅", f"已插图: {os.path.basename(img_path)}")
                return True
        rep(f"资质[{title}]", "⚠️", "文档中无对应标题段")
        return False

    insert_under_title("营业执照", find_qual("营业执照"))
    insert_under_title("特种设备生产许可证", find_qual("特种设备生产许可"))
    insert_under_title("质量管理体系证书", find_qual("质量管理体系认证"))

    # 认证表中其余资质附文档末尾，按甲方表格顺序
    extra = [
        ("环境管理体系认证证书", find_qual("环境管理体系认证")),
        ("职业健康安全管理体系认证证书", find_qual("职业健康安全管理体系认证")),
        ("建筑业企业资质证书", find_qual("建筑业企业资质")),
        ("安全生产许可证", find_qual("安全生产许可")),
    ]
    anchor = doc.paragraphs[-1]._p  # 当前文档最后一段
    for title, q in extra:
        if not q:
            rep(f"资质[{title}]", "❌", "台账无此资质")
            continue
        img_path = q.get("_resolved_path", "")
        if not img_path or not os.path.exists(img_path):
            rep(f"资质[{title}]", "❌", f"原件不存在: {img_path}")
            continue
        np1 = new_para_after(anchor)
        anchor = np1
        tp = Paragraph(np1, doc)
        r = tp.add_run(title)
        r.bold = True
        r.font.size = Pt(14)
        np2 = new_para_after(anchor)
        anchor = np2
        ip = Paragraph(np2, doc)
        irun = ip.add_run()
        try:
            irun.add_picture(img_path, width=Inches(5.0))
            rep(f"资质[{title}]", "✅", f"已附末尾: {os.path.basename(img_path)}")
        except Exception as e:
            rep(f"资质[{title}]", "❌", f"插图失败: {e}")

    # ========== 5) 文档末尾强制提示（若已有则不重复） ==========
    has_note = any(p.text.strip().startswith("【本结果由AI生成") for p in doc.paragraphs)
    if not has_note:
        np_end = new_para_after(anchor)
        np2p = Paragraph(np_end, doc)
        nr = np2p.add_run(NOTE)
        nr.bold = True
        nr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    doc.save(args.out)
    print("已保存:", args.out)
    print("\n===== 填充报告 =====")
    for pos, st, note in fill_report:
        print(f"[{st}] {pos}: {note}")


if __name__ == "__main__":
    main()
