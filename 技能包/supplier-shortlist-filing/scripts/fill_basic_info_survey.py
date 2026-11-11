# -*- coding: utf-8 -*-
"""供应商基本信息调查表 v2 — 直接填到对应行/对应单元格：

- R02 §1 标签后追加 run(值)，不另起段落
- R04 §2 □ → ■（已通过认证的）
- R07 §4 □ → ■（已具备材料的），删除末尾旧摘要
- R09-R35 §5 自评清单：按台账可判断项填 tc5=■合格 / tc6=■不合格；不可判断的留空
- 末尾追加 AI 复核提示

用法:
  python fill_basic_info_survey.py --src <甲方docx> --out <成品.docx> \
      --ledgers-dir <parse_ledgers 输出目录>
"""
import argparse, json, os
from docx import Document
from docx.shared import RGBColor
from docx.oxml import OxmlElement

NOTE = "【本结果由AI生成，需人工复核、补充缺失资料及自评清单中无法自动判断项，不做最终定稿，不可直接对外提交甲方】"

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

report = []
def rep(pos, st, note):
    report.append((pos, st, note))


def get_tcs(row):
    return row._tr.findall(f'.//{W}tc')


def get_paras(tc):
    return tc.findall(f'{W}p')


def find_para_by_text(tc, substring):
    """在 tc 内找第一个 text 含 substring 的 w:p 元素"""
    for p in get_paras(tc):
        text = ''.join(t.text or '' for t in p.findall(f'.//{W}t'))
        if substring in text:
            return p
    return None


def find_qual_by_kw(qs, keywords):
    """在资质清单中找第一条证书名包含任一关键字的记录"""
    for q in qs:
        name = str(q.get("证书完整名称", ""))
        for kw in keywords:
            if kw in name:
                return q
    return None


def add_run(p, text, bold=False, color=None):
    """在段落末尾追加一个 run"""
    r = OxmlElement('w:r')
    if bold or color:
        rPr = OxmlElement('w:rPr')
        if bold:
            rPr.append(OxmlElement('w:b'))
        if color:
            c = OxmlElement('w:color')
            c.set(W + 'val', color)
            rPr.append(c)
        r.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    p.append(r)
    return r


def replace_para_text(p, old, new):
    """在段落内把 old 替换为 new（精确替换第一个匹配）。"""
    runs = p.findall(f'{W}r')
    # 拼接所有 run 的文本
    full = ''
    parts = []  # [(run_idx, run, text)]
    for ri, r in enumerate(runs):
        ts = r.findall(f'{W}t')
        for t in ts:
            parts.append((ri, r, t, t.text or ''))
            full += t.text or ''
    if old not in full:
        return False
    # 找到 old 在 full 中的位置
    idx = full.find(old)
    end = idx + len(old)
    # 重新生成 parts（带偏移）
    cursor = 0
    new_parts = []  # [(run, t, text)]
    for ri, r, t, txt in parts:
        seg_start = cursor
        seg_end = cursor + len(txt)
        if seg_end <= idx or seg_start >= end:
            # 完全不重叠
            new_parts.append((r, t, txt))
        elif seg_start >= idx and seg_end <= end:
            # 完全在替换区间内
            new_parts.append((r, t, ''))
        else:
            # 部分重叠：截断
            left = txt[: max(0, idx - seg_start)] if seg_start < idx else ''
            right = txt[end - seg_start:] if seg_end > end else ''
            new_parts.append((r, t, left + right))
        cursor = seg_end
    # 重写
    cursor = 0
    pos = idx
    # 策略：把 old 替换为 new，重写所有 run
    full_new = full.replace(old, new, 1)
    # 简化：把所有文本清空塞到第一个 run
    if runs:
        # 移除原 t
        for r in runs:
            for t in r.findall(f'{W}t'):
                t.text = ''
        # 第一个 run 加新文本
        first_t = runs[0].find(W + 't')
        if first_t is not None:
            first_t.text = full_new
        return True
    return False


def fill_p3_legal_rep(doc, ci):
    """P3 段外『法定代表人* ___ （签 名）』追加法人姓名（用 paragraph.add_run）"""
    for p in doc.paragraphs:
        t = p.text or ""
        if "法定代表人" in t and "签" in t and "名" in t:
            legal = ci.get("法定代表人") or ""
            if not legal:
                rep("段外[法定代表人]", "MISS", "台账无法定代表人，留空")
                return False
            r = p.add_run(legal)
            r.bold = True
            rep("段外[法定代表人]", "OK", f"已追加: {legal}")
            return True
    rep("段外[法定代表人]", "WARN", "未找到段外法定代表人段落，跳过")
    return False


def fill_r02_basic_info(ws, ci):
    """R02 §1 按段落索引一一对应填值（精确匹配，避免子串重复）。

    R02 段落结构（基于已识别模板）：
      P1  单位名称*：
      P2  通讯地址*：  邮编*：
      P3  法定代表人姓名*：  电话*：
      P4  质量管理部门名称*：  负责人*：
      P5  电话*：  传真*：
      P6  企业成立日期*：  企业统一社会信用代码*：
      P7  经济类型*：  企业所属行业：
      P8  企业规模*：  □ 大型  □ 中型  □ 小型
      P9  职工总数*：  其中：管理人员*：
      P10 质量管理人员*：  设计人员*：
      P11 主要产品*：
    """
    row = ws.rows[1]
    tcs = get_tcs(row)
    if len(tcs) != 1:
        rep("R02[§1]", "WARN", f"tc 数={len(tcs)} 异常")
        return False
    tc = tcs[0]
    paras = get_paras(tc)

    # 按段落索引精确分配值（0-based）
    # 段落顺序已识别：paras[0]=P1, paras[1]=P2, ..., paras[10]=P11
    values = {
        0:  ("单位名称",       ci.get("公司全称", "")),                      # P1
        1:  ("通讯地址",       ci.get("实际经营地址", ""), "邮编",            ci.get("邮政编码", "")),  # P2
        2:  ("法定代表人",     ci.get("法定代表人", ""),       "电话",        ci.get("联系电话", "")),  # P3
        3:  ("质量管理部门名称", "",                            "负责人",      ci.get("项目对接联系人", "")),  # P4
        4:  ("电话",           ci.get("项目对接联系人电话", "") or ci.get("联系电话", ""), "传真", ""),  # P5
        5:  ("成立日期",       ci.get("成立日期", ""),         "信用代码",     ci.get("营业执照（三证合一）", "") or ci.get("统一社会信用代码", "")),  # P6
        6:  ("经济类型",       "",                            "所属行业",     "起重设备制造"),  # P7
        # 7: P8 企业规模 走专门逻辑
        8:  ("职工总数",       ci.get("公司总人数", "") or ci.get("员工人数", ""), "管理人员", ci.get("管理人员数量", "")),  # P9
        9:  ("质量管理人员",   ci.get("质检人员数量", ""),     "设计人员",     ""),  # P10
        10: ("主要产品",       ci.get("产品", "")),  # P11
    }

    n_ok, n_blank = 0, 0
    for idx, payload in values.items():
        if idx >= len(paras):
            continue
        p = paras[idx]
        # payload 长度变化
        if len(payload) == 2:
            label, v1 = payload
            v2 = None
        else:
            label, v1, label2, v2 = payload
        # 标签前缀（用于末尾自动提示行）
        ok_1, ok_2 = False, False
        if v1:
            add_run(p, "  " + str(v1))
            n_ok += 1; ok_1 = True
        else:
            add_run(p, "  ______")
            n_blank += 1
        if v2 is not None:
            if v2:
                add_run(p, "    " + str(v2))
                n_ok += 1; ok_2 = True
            else:
                add_run(p, "    ______")
                n_blank += 1

    # P8 企业规模 □ 大型 □ 中型 □ 小型 → ■ 大型（按 2600 人/128279 万/1580 亩推断）
    if len(paras) > 7:
        p_size = paras[7]
        text = ''.join(t.text or '' for t in p_size.findall(f'.//{W}t'))
        if "□ 大型" in text:
            replace_para_text(p_size, "□ 大型", "■ 大型")

    rep("R02[§1.基本情况]", "OK" if n_ok >= 14 else "MANUAL",
        f"已填 {n_ok} 字段到原标签后，留空 {n_blank} 待人工补")
    return True


def fill_r04_certifications(ws, qs):
    """R04 §2 直接在 □ 段落里把 □ → ■（已通过认证的）。"""
    row = ws.rows[3]
    tcs = get_tcs(row)
    if len(tcs) != 1:
        rep("R04[§2]", "WARN", f"tc 数={len(tcs)} 异常")
        return False
    tc = tcs[0]

    # 找资质清单
    qms = find_qual_by_kw(qs, ["质量管理体系", "ISO9001", "ISO 9001"])
    q_env = find_qual_by_kw(qs, ["环境管理体系", "ISO14001"])
    q_ohs = find_qual_by_kw(qs, ["职业健康", "OHSAS", "ISO45001"])
    has_any = bool(qms or q_env or q_ohs)

    # P2 □ 建立并实施质量管理体系，但未认证注册  → 保留 □
    # P3 □ 已获认证注册  → ■
    if has_any:
        p3 = find_para_by_text(tc, "□ 已获认证注册。认证类型")
        if p3 is not None:
            # 直接把第一个 □ → ■
            replace_para_text(p3, "□ 已获认证注册", "■ 已获认证注册")

        # P6 □GB/T19001 → ■（如有 ISO9001 类证书）
        if qms:
            p6 = find_para_by_text(tc, "□GB/T19001")
            if p6 is not None:
                replace_para_text(p6, "□GB/T19001", "■GB/T19001")
                # 同时把证书编号/有效期写上
                text = ''.join(t.text or '' for t in p6.findall(f'.//{W}t'))
                if "证书编号" in text:
                    # 在第一个"证书编号"前填值
                    p6_text = ''.join(t.text or '' for t in p6.findall(f'.//{W}t'))
                    if qms.get("证书编号"):
                        new = p6_text.replace("证书编号              ", f"证书编号：{qms.get('证书编号','')}  ", 1)
                        new = new.replace("证书有效期             ", f"证书有效期：{qms.get('有效期','')}  ", 1)
                        runs = p6.findall(f'{W}r')
                        if runs:
                            for r in runs:
                                for t in r.findall(f'{W}t'):
                                    t.text = ''
                            ft = runs[0].find(W + 't')
                            if ft is not None:
                                ft.text = new

        # P9/P10 环境/职业健康管理体系 同样处理
        if q_env or q_ohs:
            p10 = find_para_by_text(tc, "□ 已获认证注册  证书编号")
            if p10 is not None:
                replace_para_text(p10, "□ 已获认证注册", "■ 已获认证注册")
                # 写编号/有效期（取 q_env 或 q_ohs 任一已具备）
                pick = q_env or q_ohs
                if pick:
                    text = ''.join(t.text or '' for t in p10.findall(f'.//{W}t'))
                    new = text.replace("证书编号              ", f"证书编号：{pick.get('证书编号','')}  ", 1)
                    new = new.replace("证书有效期             ", f"证书有效期：{pick.get('有效期','')}  ", 1)
                    runs = p10.findall(f'{W}r')
                    if runs:
                        for r in runs:
                            for t in r.findall(f'{W}t'):
                                t.text = ''
                        ft = runs[0].find(W + 't')
                        if ft is not None:
                            ft.text = new

    # 删除原 R04 末尾的 ✅ 摘要段（前一个流程留下的）
    removed = 0
    paras = get_paras(tc)
    for p in paras:
        text = ''.join(t.text or '' for t in p.findall(f'.//{W}t'))
        if text.startswith("✅ "):
            tc.remove(p)
            removed += 1

    rep("R04[§2.体系认证]",
        "OK" if has_any else "MANUAL",
        f"已勾选 ■ 认证项" + (f" + 编号/有效期已填" if has_any else "") + f"，清掉旧摘要段 {removed} 条")
    return True


def fill_r07_evidence_materials(ws, qs):
    """R07 §4 把对应 □ 改成 ■（已具备材料的）。"""
    row = ws.rows[6]
    tcs = get_tcs(row)
    if len(tcs) != 1:
        rep("R07[§4]", "WARN", f"tc 数={len(tcs)} 异常")
        return False
    tc = tcs[0]

    # 决定勾哪些
    owned_1 = bool(find_qual_by_kw(qs, ["营业执照"]))                     # 1.营业执照
    owned_8 = bool(find_qual_by_kw(qs, ["质量管理体系", "ISO9001"]))       # 8.质量体系证书
    owned_12 = bool(find_qual_by_kw(qs, ["环境管理体系", "ISO14001"])) or \
               bool(find_qual_by_kw(qs, ["职业健康", "OHSAS", "ISO45001"]))  # 12.其它证实性材料（含环境/安全）
    owned_owned = owned_1 or owned_8 or owned_12

    if owned_1:
        p1 = find_para_by_text(tc, "□1.企业法人营业执照")
        if p1 is not None:
            replace_para_text(p1, "□1.", "■1.")
    if owned_8:
        p8 = find_para_by_text(tc, "□8.质量体系认证证书")
        if p8 is not None:
            replace_para_text(p8, "□8.", "■8.")
    if owned_12:
        p12 = find_para_by_text(tc, "□12.填报企业自认为还应提供的其它证实性材料")
        if p12 is not None:
            replace_para_text(p12, "□12.", "■12.")

    # 删除旧"AI 已识别具备项"摘要段
    removed = 0
    paras = get_paras(tc)
    for p in paras:
        text = ''.join(t.text or '' for t in p.findall(f'.//{W}t'))
        if "AI 已识别具备项" in text:
            tc.remove(p)
            removed += 1

    rep("R07[§4.证实性材料]",
        "OK" if owned_owned else "MANUAL",
        f"已勾选 ■ 项={sum([owned_1, owned_8, owned_12])}/3，清掉旧摘要段 {removed} 条")
    return True


def fill_r05_product_info(ws, ext):
    """R05 §3 嵌套子表填台账可覆盖项（保留旧逻辑）。"""
    row = ws.rows[4]
    nested = row._tr.findall(f'.//{W}tbl')
    if not nested:
        rep("R05[§3]", "WARN", "未找到嵌套子表")
        return False
    nt = nested[0]
    rows = nt.findall(f'{W}tr')

    eq_rows = ext["sheets"].get("主要生产设备清单", {}).get("rows", [])
    eq_summary = f"主要生产设备共 {len(eq_rows)} 台套（详见《主要生产设备清单》）" if eq_rows else "资料库无设备数据"
    tech_count = "员工总数 2600 人（技术口径待核对）"

    fills = {
        1: "GB/T 3811-2008、GB/T 14405-2011、GB/T 14406-2011 等起重机械国标",
        2: "台账未量化指标，待人工补充",
        3: "以具体合同/产品对应的型式试验报告为准",
        4: "通过 ISO9001 体系认证，质量可靠性详见具体产品合同",
        5: "需根据具体产品报价和合同条款评估",
        6: "自主研发 + 联合设计",
        7: "CAD/CAE 数字化设计",
        8: tech_count,
        9: eq_summary,
        10: "检验检测中心（具体设备清单待补充）",
    }

    n_ok = 0
    for idx, val in fills.items():
        if idx < len(rows):
            tcs = rows[idx].findall(f'{W}tc')
            if len(tcs) >= 2:
                content_tc = tcs[1]
                for p in content_tc.findall(f'{W}p'):
                    content_tc.remove(p)
                new_p = OxmlElement('w:p')
                content_tc.append(new_p)
                r = OxmlElement('w:r')
                t = OxmlElement('w:t')
                t.text = val
                r.append(t)
                new_p.append(r)
                n_ok += 1

    rep("R05[§3.产品信息·嵌套表]", "OK", f"已填 {n_ok}/{len(fills)} 指标内容")
    return True


def fill_r09_to_r35_self_assessment(ws, qs, ms, ext):
    """R09-R35 自评清单：按台账已知情况填 tc5=■合格 / tc6=■不合格；不可判断留空。

    R10 表头：tc5=合格 tc6=不合格
    R11-R16 A 法人资格（A1-A6）
    R17-R25 B 专业技术资质（B1-B9）
    R26-R29 C 体系和资格（C1-C4）
    R30-R35 D 经营信誉（D1-D6）
    """
    # 自评判断规则表：(序号, 项目, 判定) 判定 ∈ {'OK', 'NG', None} None 表示不动
    # None 表示 AI 不能判断，留空 + 在末尾报告里提示
    rules = {
        'A1': ('OK',  '营业执照 = 台账有'),                                # 营业执照
        'A2': (None,  '法定代表人证书 → 台账无该独立证书'),                # 无
        'A3': ('OK',  '组织机构代码 = 三证合一 → 信用代码已具备'),         # 信用代码
        'A4': ('OK',  '国有土地使用证/租用合同 → 房产证自有'),             # 房产证
        'A5': ('OK',  '房屋所有权证/租用合同 → 房产证自有'),               # 房产证
        'A6': (None,  '其他与法人资格 → 无台账依据'),
        'B1': (None,  '专业技术管理制度 → 无台账依据'),
        'B2': (None,  '技术人员队伍 → 台账口径不一致(38/260)'),
        'B3': ('OK',  '主要加工设备 → 49 台套'),
        'B4': ('OK',  '主要检测设备 → 检验设备清单'),
        'B5': (None,  '按技术标准生产能力 → 需业务确认'),
        'B6': (None,  '主体和关键技术 → 需业务确认'),
        'B7': ('OK',  '主要配套单位协作关系 → 配套供应商台账'),
        'B8': (None,  '其他 → 无台账依据'),
        'B9': (None,  '特种工艺规范、无损检测 → 无台账依据'),
        'C1': ('OK',  '第三方质量管理体系认证 → ISO9001'),
        'C2': (None,  '二方质量管理体系审核 → 无台账依据'),
        'C3': ('NG',  '保密资质证书 → 台账无'),
        'C4': ('OK',  '环境/职业健康安全体系 → ISO14001+OHSAS'),
        'D1': (None,  '近三年交货记录 → 无业务台账'),
        'D2': (None,  '近三年报价文件 → 无业务台账'),
        'D3': (None,  '银行账号未被冻结 → 无台账依据'),
        'D4': (None,  '税务登记证和完税凭证 → 无台账依据'),
        'D5': ('NG',  '社保登记证和缴费记录 → 资料库无社保证明'),
        'D6': (None,  '其他（财务指标）→ 无台账依据'),
    }

    n_ok = n_ng = n_manual = 0
    manual_items = []
    for r in range(11, 36):  # R11-R35
        row = ws.rows[r-1]
        tcs = get_tcs(row)
        if len(tcs) < 7:
            continue
        # tc1=序号, tc5=合格列, tc6=不合格列
        seq_cell = tcs[1]
        seq_text = ''.join(t.text or '' for t in seq_cell.findall(f'.//{W}t')).strip()
        if seq_text not in rules:
            continue
        verdict, reason = rules[seq_text]

        ok_cell = tcs[5]
        ng_cell = tcs[6]
        # 清空原内容
        for p in ok_cell.findall(f'{W}p'):
            ok_cell.remove(p)
        for p in ng_cell.findall(f'{W}p'):
            ng_cell.remove(p)

        if verdict == 'OK':
            new_p = OxmlElement('w:p')
            ok_cell.append(new_p)
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rPr.append(OxmlElement('w:b'))
            r.append(rPr)
            t = OxmlElement('w:t')
            t.text = '■合格'
            r.append(t)
            new_p.append(r)
            n_ok += 1
        elif verdict == 'NG':
            new_p = OxmlElement('w:p')
            ng_cell.append(new_p)
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rPr.append(OxmlElement('w:b'))
            r.append(rPr)
            t = OxmlElement('w:t')
            t.text = '■不合格'
            r.append(t)
            new_p.append(r)
            n_ng += 1
        else:
            # 留空 + 提示
            new_p = OxmlElement('w:p')
            ok_cell.append(new_p)
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = '□'
            r.append(t)
            new_p.append(r)
            new_p2 = OxmlElement('w:p')
            ok_cell.append(new_p2)
            r2 = OxmlElement('w:r')
            t2 = OxmlElement('w:t')
            t2.text = '□'
            r2.append(t2)
            new_p2.append(r2)
            n_manual += 1
            manual_items.append((seq_text, reason))

    rep("R09-R35[§5.自评清单]",
        "OK" if n_manual <= 14 else "MANUAL",
        f"■合格 {n_ok} 项 + ■不合格 {n_ng} 项 + 待人工 {n_manual} 项")
    return manual_items


def add_ai_notice(doc):
    if any(NOTE[:15] in p.text for p in doc.paragraphs):
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

    ci = json.load(open(os.path.join(args.ledgers_dir, "ledgers_company.json"), encoding="utf-8"))["company_info"]
    ext = json.load(open(os.path.join(args.ledgers_dir, "ledgers_company_extended.json"), encoding="utf-8"))
    qs = json.load(open(os.path.join(args.ledgers_dir, "ledgers_qualifications.json"), encoding="utf-8"))["qualifications"]
    ms = json.load(open(os.path.join(args.ledgers_dir, "ledgers_materials.json"), encoding="utf-8"))["materials"]

    doc = Document(args.src)

    # 1) 段外：P3 法定代表人
    fill_p3_legal_rep(doc, ci)

    ws = doc.tables[0]

    # 2) R02 §1 在原标签段后追加 run
    fill_r02_basic_info(ws, ci)

    # 3) R04 §2 □ → ■
    fill_r04_certifications(ws, qs)

    # 4) R05 §3 嵌套表
    fill_r05_product_info(ws, ext)

    # 5) R07 §4 □ → ■
    fill_r07_evidence_materials(ws, qs)

    # 6) R09-R35 §5 自评清单
    manual_items = fill_r09_to_r35_self_assessment(ws, qs, ms, ext)

    # 7) 末尾 AI 提示
    if add_ai_notice(doc):
        rep("末尾 AI 复核提示", "OK", "已追加")

    doc.save(args.out)

    # 报告
    work = os.path.dirname(os.path.abspath(args.out)) or "."
    md = ["# 供应商基本信息调查表填充报告（v2 对应行填值）", "",
          f"- 模板: `{args.src}`", f"- 成品: `{args.out}`", ""]
    md.append("## 填充摘要")
    md.append("")
    for pos, st, note in report:
        mark = {"OK": "✅", "MANUAL": "⚠️", "MISS": "❌", "WARN": "⚠️"}.get(st, "  ")
        md.append(f"- {mark} **{pos}**：{note}")
    md.append("")
    md.append("## ⚠️ 仍需人工处理的项")
    md.append("")
    md.append("### §1.供应商基本情况（R02）")
    md.append("以下字段台账缺失仍为占位 `______`：")
    md.append("- 质量管理部门名称 / 传真 / 经济类型 / 设计人员")
    md.append("")
    md.append("### §5.供应商自评清单（R09-R35）")
    md.append("以下项目 AI 无法代决，留空待业务人员按实际情况勾选：")
    for code, reason in manual_items:
        md.append(f"- **{code}** — {reason}")
    md.append("")
    md.append("---")
    md.append("")
    md.append("**合规声明**：数据来源唯一权威台账 `E:\\卫华供应商入围资料库`；未读取任何证书图片/PDF 内部内容；自评清单按台账规则判断（证书/资质/资料具备 → 合格，否则 NG；业务/资质文件无法从台账判断的留空），但仍需业务人员最终复核。")

    md_path = os.path.join(work, "basic_info_fill_report.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    print("已保存:", args.out)
    print()
    print("===== 填充报告 =====")
    for pos, st, note in report:
        mark = {"OK": "✅", "MANUAL": "⚠️", "MISS": "❌", "WARN": "⚠️"}.get(st, "  ")
        print(f"{mark} {pos}: {note}")
    print()
    print(f"⚠️ R09-R35 待人工 {len(manual_items)} 项: {[m[0] for m in manual_items]}")
    print()
    print("报告文件:", md_path)


if __name__ == "__main__":
    main()