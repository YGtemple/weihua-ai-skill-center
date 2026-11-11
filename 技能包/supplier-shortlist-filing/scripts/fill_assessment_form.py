# -*- coding: utf-8 -*-
"""供应商考核/评鉴相关表单填写（CLI 版）：
主表填值 + 勾选，4 个附件 sheet 插入对应原件。

用法:
  python fill_assessment_form.py --src <甲方模板.xlsx> --out <成品.xlsx> \
      --ledgers-dir <parse_ledgers 输出目录>

数据源规则（与 skill 一致）:
  - 公司信息取 ledgers_company.json 的 company_info
  - 客户群/生产设备取 ledgers_company_extended.json 的多工作表数据
  - 附件原件一律按台账 _resolved_path 定位，禁止硬编码猜测
  - 台账无值/缺失 → 留空并标 ⚠️/❌，禁止编造
"""
import argparse, json, os, sys
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment

NOTE = "【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】"

report = []
def rep(pos, st, note):
    report.append((pos, st, note))

def merged_top_left(ws, coord):
    """若 coord 落在合并区，返回该区左上角坐标；否则原样返回。"""
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
    row, col = coordinate_to_tuple(coord)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return f"{get_column_letter(rng.min_col)}{rng.min_row}"
    return coord

def set_ws_value(ws, coord, val, wrap=False):
    """写值（自动落到合并区左上角）；长文本可选自动换行。"""
    c = merged_top_left(ws, coord)
    ws[c] = val
    if wrap:
        ws[c].alignment = Alignment(wrap_text=True, vertical="center")
    return c

def find_qualification(qs, keyword):
    """在资质清单中按证书名关键字找记录，返回 _resolved_path 或 None。"""
    for q in qs:
        if keyword in str(q.get("证书完整名称", "")) or keyword in str(q.get("类别", "")):
            p = q.get("_resolved_path")
            if p and os.path.exists(p):
                return p
    return None

def find_material(ms, keyword):
    """在资料清单中按文件名关键字找记录，返回 _resolved_path 或 None。"""
    for m in ms:
        if keyword in str(m.get("文件名", "")):
            p = m.get("_resolved_path")
            if p and os.path.exists(p):
                return p
    return None

def pdf_to_pngs(pdf_path, out_dir, prefix, max_pages):
    """PDF 页面转 png（仅格式转换用于附件放置，绝不解析内容）。"""
    import pymupdf
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    paths = []
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        pix = page.get_pixmap(dpi=110)
        p = os.path.join(out_dir, f"{prefix}_p{i+1}.png")
        pix.save(p)
        paths.append(p)
    doc.close()
    return paths

def insert_image(ws, path, anchor, scale=0.5):
    ximg = XLImage(path)
    ximg.width, ximg.height = int(ximg.width * scale), int(ximg.height * scale)
    ws.add_image(ximg, anchor)
    return ximg

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="甲方模板 xlsx")
    ap.add_argument("--out", required=True, help="成品输出 xlsx")
    ap.add_argument("--ledgers-dir", required=True, help="parse_ledgers 输出目录")
    args = ap.parse_args()

    work = os.path.dirname(os.path.abspath(args.out)) or "."
    ci = json.load(open(os.path.join(args.ledgers_dir, "ledgers_company.json"), encoding="utf-8"))["company_info"]
    ext = json.load(open(os.path.join(args.ledgers_dir, "ledgers_company_extended.json"), encoding="utf-8"))
    qs = json.load(open(os.path.join(args.ledgers_dir, "ledgers_qualifications.json"), encoding="utf-8"))["qualifications"]
    ms = json.load(open(os.path.join(args.ledgers_dir, "ledgers_materials.json"), encoding="utf-8"))["materials"]
    def g(k): return ci.get(k, "") or ""

    wb = openpyxl.load_workbook(args.src)
    ws = wb["03-供应商评鉴表--资料评鉴"]

    # ===== 1) 主表基础信息 =====
    fills = [
        ("C4", "生产商/加工类", "开发类别（模板预设，确认即可）"),
        ("C5", g("公司全称"), "台账键:公司全称"),
        ("C6", g("实际经营地址"), "台账键:实际经营地址"),
        ("C7", g("法定代表人"), "台账键:法定代表人"),
        ("F7", g("项目对接联系人"), "台账键:项目对接联系人"),
        ("C8", g("联系电话"), "台账键:联系电话"),
    ]
    for addr, val, note in fills:
        set_ws_value(ws, addr, val)
        rep(f"评鉴表!{addr}", "OK" if val else "MANUAL", note + (" = " + str(val) if val else " 台账无值留空"))

    # 传真：台账无 → 留空
    set_ws_value(ws, "F8", "")
    rep("评鉴表!F8 传真", "MANUAL", "台账无传真号，留空待人工补充")

    # ===== 2) 厂商概况（含历史修复项：员工人数/上年度销售额现在可正确读取） =====
    overview = [
        ("D10", g("成立日期"), "台账键:成立日期"),
        ("G10", g("员工人数") or g("公司总人数"), "台账键:员工人数/公司总人数（日期守卫修复后可正常读取）"),
        ("D11", g("上年度销售额"), "台账键:上年度销售额"),
        ("G11", f"{g('注册资本')} / {g('实缴资金')}", "台账键:注册资本+实缴资金"),
        ("D12", g("生产/服务场地面积"), "台账键:生产/服务场地面积"),
        ("G12", "", "台账无库房面积键值，留空待人工补充"),
        ("D13", g("产品") + "；" + str(g("经营范围"))[:120], "台账键:产品+经营范围"),
    ]
    for addr, val, note in overview:
        set_ws_value(ws, addr, val, wrap=True)
        st = "OK" if val else "MANUAL"
        rep(f"评鉴表!{addr}", st, note + (" = " + str(val)[:60] if val else " 留空待人工"))

    # ===== 3) 客户群 / 生产设备（多工作表规则：台账《供方能力调查基础信息表》子表） =====
    cust_rows = ext["sheets"].get("主要客户信息", {}).get("rows", [])
    if cust_rows:
        names = [r.get("主要客户信息", "") for r in cust_rows if r.get("主要客户信息")]
        txt = "、".join(names[:12]) + f"（共{len(names)}家，详见《主要客户信息》）" if len(names) > 12 else "、".join(names)
        set_ws_value(ws, "D14", txt, wrap=True)
        rep("评鉴表!D14 客户群", "OK", f"已填 {len(names)} 家客户名称（展示前12家+总数）")
    else:
        rep("评鉴表!D14 客户群", "MANUAL", "台账无客户群明细，留空待人工补充")

    eq_rows = ext["sheets"].get("主要生产设备清单", {}).get("rows", [])
    if eq_rows:
        desc = []
        for r in eq_rows[:8]:
            name = r.get("设备名称", "")
            maker = r.get("制造厂家", "")
            desc.append(f"{name}（{maker}）" if maker else name)
        txt = "、".join(desc) + f"（共{len(eq_rows)}台套，详见《主要生产设备清单》）" if len(eq_rows) > 8 else "、".join(desc)
        set_ws_value(ws, "D15", txt, wrap=True)
        rep("评鉴表!D15 生产设备", "OK", f"已填 {len(eq_rows)} 台套主要设备（展示前8类+总数）")
    else:
        rep("评鉴表!D15 生产设备", "MANUAL", "台账无设备明细，留空待人工补充")

    # ===== 4) A3/A9 勾选 =====
    a3 = str(ws["A3"].value or "")
    if "■ 生产商/加工类" not in a3:
        ws["A3"] = a3.replace("□ 生产商/加工类", "■ 生产商/加工类")
        rep("评鉴表!A3", "OK", "已勾选 生产商/加工类")
    else:
        rep("评鉴表!A3", "OK", "模板已勾选 生产商/加工类")

    a9 = str(ws["A9"].value or "")
    new9 = (a9.replace("□营业执照", "■营业执照")
              .replace("□ 营业执照", "■ 营业执照")
              .replace("□资质证书", "■资质证书")
              .replace("□ 资质证书", "■ 资质证书")
              .replace("□ISO9000认证", "■ISO9000认证")
              .replace("□ISO9000", "■ISO9000")
              .replace("□ ISO9000", "■ ISO9000"))
    has_any_opt = any(k in a9 for k in ("□营业执照", "□ 营业执照", "□资质证书", "□ISO9000", "□代理合同", "□食品"))
    if new9 != a9:
        ws["A9"] = new9
        rep("评鉴表!A9 证照勾选", "OK", "已勾:营业执照/资质证书/ISO9000认证")
    elif not has_any_opt:
        rep("评鉴表!A9", "MANUAL", "模板 A9 未见 □勾选项文本，请人工核对证照勾选")
    else:
        rep("评鉴表!A9", "OK", "证照勾选项已就绪（无需改动）")

    # ===== 5) 末尾 AI 提示（防重复） =====
    existing_note = any(
        str(ws.cell(row=r, column=1).value or "").startswith(NOTE[:15])
        for r in range(ws.max_row - 2, ws.max_row + 1)
    )
    if not existing_note:
        ws["A30"] = NOTE
        ws["A30"].font = Font(bold=True, color="C00000", size=10)
        ws["A30"].alignment = Alignment(wrap_text=True)
        ws.merge_cells("A30:H31")
        rep("评鉴表!A30", "OK", "末尾已追加 AI 复核提示")
    else:
        rep("评鉴表!A30", "OK", "检测到已有 AI 提示，未重复追加")

    # ===== 6) 附件 sheet（原件一律按台账 _resolved_path 定位） =====
    tmp_dir = os.path.join(work, "_tmp_att")
    os.makedirs(tmp_dir, exist_ok=True)

    # 6.1 营业执照
    ws_bl = wb["营业执照"]
    ws_bl["A1"] = "营业执照（原件扫描）"
    ws_bl["A1"].font = Font(bold=True, size=12)
    img_lic = find_qualification(qs, "营业执照")
    if img_lic:
        insert_image(ws_bl, img_lic, "A3", scale=0.55)
        rep("附件页[营业执照]", "OK", f"已插入原件扫描图（{os.path.basename(img_lic)}）")
    else:
        ws_bl["A2"] = "❌资料库缺失：营业执照原件未找到"
        rep("附件页[营业执照]", "MISS", "台账中未定位到营业执照文件")

    # 6.2 社保证明
    ws_sb = wb["社保证明"]
    pdf_sb = find_material(ms, "社保")
    if pdf_sb:
        try:
            pngs = pdf_to_pngs(pdf_sb, tmp_dir, "sb", 6)
            ws_sb["A1"] = "社保证明（原件扫描）"
            ws_sb["A1"].font = Font(bold=True, size=12)
            row = 3
            for p in pngs:
                insert_image(ws_sb, p, f"A{row}", scale=0.5)
                row += 42
            rep("附件页[社保证明]", "OK", f"已插入社保证明转图（{len(pngs)}页）")
        except Exception as e:
            ws_sb["A1"] = f"⚠️社保证明转图失败，需人工插入：{e}"
            rep("附件页[社保证明]", "MANUAL", f"转图失败: {e}")
    else:
        ws_sb["A1"] = "❌资料库缺失：资料库内无社保证明文件，需人工补充（台账《供应商入围资料清单》检索'社保'无记录）"
        ws_sb["A1"].font = Font(bold=True, color="C00000")
        rep("附件页[社保证明]", "MISS", "台账检索'社保'0记录，资料库缺失")

    # 6.3 其他合作厂商合作合同（PDF 转图，仅格式转换放置）
    ws_ht = wb["其他合作厂商合作合同"]
    ws_ht["A1"] = "其他合作厂商合作合同（附件原件）"
    ws_ht["A1"].font = Font(bold=True, size=12)
    ht_pdfs = [
        ("近三年合同履约情况自证性文件.pdf", find_material(ms, "合同履约")),
        ("服务进度、质量、交付满足合同.pdf", find_material(ms, "服务进度、质量、交付满足合同")),
    ]
    row = 3
    n_pages = 0
    for name, path in ht_pdfs:
        if not path:
            ws_ht.cell(row=row, column=1, value=f"❌缺失: {name}")
            ws_ht.cell(row=row, column=1).font = Font(bold=True, color="C00000")
            row += 1
            rep(f"附件页[合作合同-{name}]", "MISS", "台账未定位到该文件")
            continue
        pngs = pdf_to_pngs(path, tmp_dir, "ht_" + name[:4], 4)
        ws_ht.cell(row=row, column=1, value=f"◆ {name}（共{min(4, len(pngs))}页，超出4页部分请对照原件补充）").font = Font(bold=True)
        row += 1
        for p in pngs:
            insert_image(ws_ht, p, f"A{row}", scale=0.5)
            row += 42
        n_pages += len(pngs)
    rep("附件页[合作合同]", "OK" if n_pages else "MANUAL", f"合同 PDF 转图共 {n_pages} 页（格式转换放置，未解析内容）")

    # 6.4 房产证
    ws_fc = wb["房屋租赁或购买合同（房产证）"]
    ws_fc["A1"] = "房屋租赁或购买合同（房产证，附件原件）"
    ws_fc["A1"].font = Font(bold=True, size=12)
    pdf_fc = find_material(ms, "房产证")
    if pdf_fc:
        try:
            pngs = pdf_to_pngs(pdf_fc, tmp_dir, "fc", 6)
            row = 3
            for p in pngs:
                insert_image(ws_fc, p, f"A{row}", scale=0.5)
                row += 42
            rep("附件页[房产证]", "OK", f"已插入房产证.pdf 转图（{len(pngs)}页）")
        except Exception as e:
            ws_fc["A2"] = f"⚠️房产证转图失败，需人工插入：{e}"
            rep("附件页[房产证]", "MANUAL", f"转图失败: {e}")
    else:
        ws_fc["A2"] = "❌缺失：房产证.pdf 台账未定位到"
        rep("附件页[房产证]", "MISS", "台账检索'房产证'未定位文件")

    wb.save(args.out)

    # ===== 7) 填充报告 =====
    md = ["# 供应商评鉴表填充报告", "", f"- 模板: `{args.src}`", f"- 成品: `{args.out}`", ""]
    for pos, st, note in report:
        md.append(f"- {'✅' if st=='OK' else '⚠️' if st=='MANUAL' else '❌'} **{pos}**：{note}")
    md_path = os.path.join(work, "assessment_fill_report.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    print("已保存:", args.out)
    print()
    print("===== 填充报告 =====")
    for pos, st, note in report:
        mark = {"OK": "✅", "MANUAL": "⚠️", "MISS": "❌"}.get(st, "  ")
        print(f"{mark} {pos}: {note}")
    print()
    print("报告文件:", md_path)

if __name__ == "__main__":
    main()
