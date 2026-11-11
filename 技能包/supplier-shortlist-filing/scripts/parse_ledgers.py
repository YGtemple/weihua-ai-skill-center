# -*- coding: utf-8 -*-
"""解析 E:\\卫华供应商入围资料库 下三份台账，输出结构化 JSON + 文件存在性校验报告。

用法:
    python parse_ledgers.py --out <输出目录> [--base E:\\卫华供应商入围资料库]

输出:
    <out>/ledgers_company.json              公司基础信息（仅 Sheet1 项目-内容键值对）
    <out>/ledgers_company_extended.json     多工作表全集（公司+客户+生产设备+检验设备+
                                            配套外协供应商+设备年产量+财务信息）— 2026-09 追加
    <out>/ledgers_qualifications.json       资质清单
    <out>/ledgers_materials.json            供应商入围资料清单（含实际文件定位结果）
    <out>/ledgers_verify_report.md          校验报告（路径修复、缺失、告警）

硬性规则:
    1) 三份 Excel 台账是唯一权威数据源；禁止读取证书图片/PDF 内部内容。
    2) 解析失败/结构异常 → 明确告警退出，禁止编造。
    3) 多工作表解析：每个 sheet 独立识别，按表头逐行读取，禁止只认两列键值对，
       禁止跳过非键值表（如主要客户信息、主要生产设备清单等）。
"""
import argparse
import json
import os
import sys

# 允许从 scripts 目录外直接运行
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl  # noqa: E402

from common import (  # noqa: E402
    DEFAULT_BASE, LEDGER_FILES, MARK_MISSING, cell_str, cell_str_legacy, die, ensure_dir,
    excel_serial_to_date, resolve_file,
)


def load_wb(base, key):
    path = os.path.join(base, LEDGER_FILES[key])
    if not os.path.isfile(path):
        die(f"台账缺失: {path}。请检查 E:\\卫华供应商入围资料库 目录，禁止编造数据。")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        die(f"台账读取失败 {path}: {e}")
    return wb, path


def parse_company(base):
    """仅解析 Sheet1『供方能力调查基础信息』两列键值对（保持向后兼容）。
    使用 cell_str_legacy 让成立日期=39120 仍能 → '2007-02-07'。"""
    wb, path = load_wb(base, "company")
    ws = None
    for name in wb.sheetnames:
        if "基础信息" in name:
            ws = wb[name]
            break
    if ws is None:
        die(f"《供方能力调查基础信息表》未找到含『基础信息』的 sheet: {wb.sheetnames}")
    pairs = {}
    for row in ws.iter_rows(values_only=True):
        if row is None or len(row) < 2:
            continue
        key, val = row[0], row[1]
        key_s = cell_str_legacy(key)
        if not key_s or key_s in ("项目/字段名称", "项目-内容", "项目", "填写内容"):
            continue
        val_s = cell_str_legacy(val)
        if not val_s:
            continue
        pairs[key_s] = val_s
    if not pairs:
        die(f"《供方能力调查基础信息表》解析结果为空，结构异常（预期 A列=项目 B列=内容）。请人工核对台账，禁止编造。")
    return {"source_file": os.path.basename(path), "company_info": pairs}


# ---------------------------------------------------------------------------
# 多工作表专项解析（2026-09 追加规则）
# ---------------------------------------------------------------------------

# 7 个工作表识别规则：sheet_name 关键字 → 类型标识
SHEET_RULES = [
    # (sheet_name 包含的关键字, 类型标识, 中文说明)
    ("基础信息",       "basic_info",    "供方能力调查基础信息（项目-内容键值对）"),
    ("主要客户信息",   "customers",       "主要客户信息列表"),
    ("主要生产设备",   "production_eq",   "主要生产设备清单"),
    ("主要检验设备",   "inspection_eq",   "主要检验设备清单"),
    ("配套件供应商",   "suppliers",       "配套件/外协件供应商"),
    ("设备年产量",     "annual_capacity", "设备年产量"),
    ("财务信息",       "financial",       "财务信息（横向年度表）"),
]


def _find_sheet(wb, keyword):
    """按关键字匹配工作表名（含模糊匹配）；找不到返回 None。"""
    for name in wb.sheetnames:
        if keyword in name:
            return wb[name], name
    return None, None


def _parse_table_sheet(ws):
    """通用表格型 sheet：第一行为表头，逐行读取为 list of dict。
    自动跳过全空行；保留行原顺序。
    按 cell.number_format 决定是否把数值当日期，避免把'使用量 9656'误转日期。"""
    rows_iter = list(ws.iter_rows(values_only=False))
    if not rows_iter:
        return [], []
    header_cells = rows_iter[0]
    header = [cell_str(c.value, c.number_format) for c in header_cells]
    # 取最后一个非空表头列下标
    last_idx = 0
    for i, c in enumerate(header_cells):
        if cell_str(c.value, c.number_format):
            last_idx = i
    header = header[: last_idx + 1]
    records = []
    for row in rows_iter[1:]:
        if len(row) < len(header):
            row = list(row) + [None] * (len(header) - len(row))
        vals = [cell_str(c.value, c.number_format) for c in row[: len(header)]]
        if not any(vals):
            continue
        rec = dict(zip(header, vals))
        records.append(rec)
    return header, records


def _parse_financial_sheet(ws):
    """财务信息：横向年度表（第一列指标名，其余列为年度）。
    输出: {'years': [...], 'metrics': [{'name': str, 'values': {年度: 值, ...}}]}"""
    rows_iter = list(ws.iter_rows(values_only=False))
    if not rows_iter:
        return {"years": [], "metrics": []}
    header_cells = rows_iter[0]
    header = [cell_str(c.value, c.number_format) for c in header_cells]
    years = header[1:]
    metrics = []
    for row in rows_iter[1:]:
        cells = list(row)
        name = cell_str(cells[0].value, cells[0].number_format) if cells else ""
        if not name:
            continue
        values = {}
        for i, y in enumerate(years):
            if i + 1 < len(cells):
                    v = cell_str(cells[i + 1].value, cells[i + 1].number_format)
                    if v:
                        values[y] = v
        if values:
            metrics.append({"name": name, "values": values})
    return {"years": years, "metrics": metrics}


def _parse_annual_capacity_sheet(ws):
    """设备年产量：每行 6 列，按 (label, value, label, value, label, value) 配对，
    提取为 [{序号, 品名, 规格, 年产量}, ...]。

    行结构示例:
        ['1.品名', '单梁桥式起重机', '规格', 'LD32T及以下', '年产量', '28800台']
        ['2.品名', '双梁桥式起重机', '规格', 'QD300及以下', '年产量', '3000台']
    """
    rows_iter = list(ws.iter_rows(values_only=False))
    if not rows_iter:
        return []
    records = []
    for row in rows_iter:
        cells = list(row)
        if len(cells) < 6:
            cells = cells + [None] * (6 - len(cells))
        label1 = cell_str(cells[0].value, cells[0].number_format)
        val1   = cell_str(cells[1].value, cells[1].number_format)
        label2 = cell_str(cells[2].value, cells[2].number_format)
        val2   = cell_str(cells[3].value, cells[3].number_format)
        label3 = cell_str(cells[4].value, cells[4].number_format)
        val3   = cell_str(cells[5].value, cells[5].number_format)
        if not (label1 and val1):
            continue
        # 序号从 label1 中提取（如 "1.品名" → "1"）
        seq = ""
        for ch in label1:
            if ch.isdigit():
                seq += ch
            else:
                break
        records.append({
            "序号": seq or str(len(records) + 1),
            "品名": val1,
            "规格": val2,
            "年产量": val3,
        })
    return records


def parse_company_extended(base):
    """读取《供方能力调查基础信息表》全部 7 个 sheet。

    输出结构:
    {
        "source_file": "供方能力调查基础信息表.xlsx",
        "company_info": {键值对},
        "sheets": {
            "供方能力调查基础信息":   {type: 'basic_info', header: [...], rows: [...]},
            "主要客户信息":           {type: 'customers',   header: [...], rows: [...]},
            "主要生产设备清单":       {type: 'production_eq', header:[...], rows: [...]},
            "主要检验设备清单":       {type: 'inspection_eq', header:[...], rows: [...]},
            "配套件供应商、外协件供应商": {type:'suppliers',   header:[...], rows: [...]},
            "设备年产量":             {type:'annual_capacity', rows: [{序号,品名,规格,年产量},...]},
            "财务信息":               {type:'financial',   years:[...], metrics:[{name,values},...]}
        },
        "missing_sheets": [未识别的 sheet 名]
    }
    """
    wb, path = load_wb(base, "company")

    # 1) Sheet1 项目-内容（与 parse_company 保持完全一致，使用 legacy 兼容 39120→2007-02-07）
    pairs = {}
    ws1 = None
    for name in wb.sheetnames:
        if "基础信息" in name:
            ws1 = wb[name]
            break
    if ws1 is not None:
        for row in ws1.iter_rows(values_only=True):
            if row is None or len(row) < 2:
                continue
            key, val = row[0], row[1]
            key_s = cell_str_legacy(key)
            if not key_s or key_s in ("项目/字段名称", "项目-内容", "项目", "填写内容"):
                continue
            val_s = cell_str_legacy(val)
            if val_s:
                pairs[key_s] = val_s

    # 2) 多工作表识别
    sheets_data = {}
    missing = []
    matched_sheet_names = set()
    for keyword, type_id, desc in [(k, t, d) for k, t, d in SHEET_RULES]:
        ws, real_name = _find_sheet(wb, keyword)
        if ws is None:
            missing.append(f"{keyword}（{desc}）")
            continue
        matched_sheet_names.add(real_name)
        if type_id == "financial":
            sheets_data[real_name] = {"type": type_id, **_parse_financial_sheet(ws)}
        elif type_id == "basic_info":
            # 已在 pairs 里；只在这里记元数据
            sheets_data[real_name] = {"type": type_id, "key_count": len(pairs)}
        elif type_id == "annual_capacity":
            sheets_data[real_name] = {"type": type_id, "rows": _parse_annual_capacity_sheet(ws)}
        else:
            header, records = _parse_table_sheet(ws)
            sheets_data[real_name] = {"type": type_id, "header": header, "rows": records}

    # 记录未匹配到的额外 sheet
    extra = [n for n in wb.sheetnames if n not in matched_sheet_names]
    if extra:
        missing.extend([f"{n}（未识别的额外 sheet）" for n in extra])

    if not sheets_data:
        die(f"《供方能力调查基础信息表》多工作表解析全部失败: {wb.sheetnames}")

    return {
        "source_file": os.path.basename(path),
        "company_info": pairs,
        "sheets": sheets_data,
        "missing_sheets": missing,
    }


def parse_qualifications(base):
    """读取《资质清单》。按 cell.number_format 决定是否转日期，避免'证书编号'数字被误转日期。"""
    wb, path = load_wb(base, "qualification")
    ws = wb[wb.sheetnames[0]]
    rows = []
    header_cells = next(ws.iter_rows(values_only=False), None)
    if header_cells is None:
        die(f"《资质清单》无表头行")
    header = [cell_str_legacy(c.value) for c in header_cells]
    for row in ws.iter_rows(values_only=False):
        rec = {}
        for c, h in zip(row, header):
            rec[h] = cell_str(c.value, c.number_format)
        if not any(v for v in rec.values()):
            continue
        if rec.get("类别") == "类别" or rec.get("证书完整名称") == "证书完整名称":
            continue  # 表头行
        if not rec.get("类别") and not rec.get("证书完整名称"):
            continue
        # 资质文件定位：先按台账路径，失败则用路径末段文件名递归搜索兜底
        rel = rec.get("文件路径（资料库内路径）") or ""
        fname_hint = os.path.basename(rel.replace("/", "\\")).strip() if rel else None
        full, note = resolve_file(base, rel, fname_hint)
        rec["_resolved_path"] = full
        rec["_resolve_note"] = note
        rows.append(rec)
    if not rows:
        die(f"《资质清单》解析结果为空，结构异常。请人工核对台账，禁止编造。")
    return {"source_file": os.path.basename(path), "qualifications": rows}


def parse_materials(base):
    """读取《供应商入围资料清单》。"""
    wb, path = load_wb(base, "material")
    ws = wb[wb.sheetnames[0]]
    rows = []
    header_cells = next(ws.iter_rows(values_only=False), None)
    if header_cells is None:
        die(f"《供应商入围资料清单》无表头行")
    header = [cell_str_legacy(c.value) for c in header_cells]
    for row in ws.iter_rows(values_only=False):
        rec = {}
        for c, h in zip(row, header):
            rec[h] = cell_str(c.value, c.number_format)
        if not any(v for v in rec.values()):
            continue
        fn = rec.get("文件名") or ""
        lp = rec.get("文件完整路径") or ""
        if fn == "文件名" or lp == "文件完整路径":
            continue  # 表头行
        if not fn and not lp:
            continue
        full, note = resolve_file(base, lp, fn)
        rec["_resolved_path"] = full
        rec["_resolve_note"] = note
        rows.append(rec)
    if not rows:
        die(f"《供应商入围资料清单》解析结果为空，结构异常。请人工核对台账，禁止编造。")
    return {"source_file": os.path.basename(path), "materials": rows}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=DEFAULT_BASE)
    ap.add_argument("--out", required=True, help="输出目录（副本目录，禁止写入源目录）")
    ap.add_argument("--skip-extended", action="store_true",
                    help="跳过解析供方能力调查基础信息表的多工作表（仅产出 ledgers_company.json）")
    args = ap.parse_args()

    if not os.path.isdir(args.base):
        die(f"资料库目录不存在: {args.base}。请确认已授权访问，禁止编造任何信息。")
    out = ensure_dir(args.out)

    company = parse_company(args.base)
    quals = parse_qualifications(args.base)
    mats = parse_materials(args.base)

    if args.skip_extended:
        company_ext = None
    else:
        company_ext = parse_company_extended(args.base)

    # 校验报告
    report = ["# 台账解析校验报告", ""]
    report.append(f"- 数据源目录: {args.base}")
    report.append(f"- 生成时间: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("")
    report.append("## 公司基础信息（Sheet1 项目-内容）")
    report.append(f"- 键值对数量: {len(company['company_info'])}（来源: {company['source_file']}）")
    report.append("")

    if company_ext is not None:
        report.append("## 供方能力调查基础信息表-多工作表（2026-09 追加）")
        for real_name, data in company_ext["sheets"].items():
            if data["type"] == "basic_info":
                report.append(f"- 【供方能力调查基础信息】键值对数: {data.get('key_count', 0)}")
            elif data["type"] == "financial":
                n_metrics = len(data.get("metrics", []))
                years = "、".join(data.get("years", []))
                report.append(f"- 【财务信息】指标数: {n_metrics} | 年度: {years}")
            else:
                rows = data.get("rows", [])
                hdr = data.get("header", [])
                report.append(f"- 【{real_name}】行数: {len(rows)} | 表头: {' / '.join(hdr)}")
        if company_ext.get("missing_sheets"):
            report.append("")
            report.append("### ⚠️ 未匹配的工作表")
            for m in company_ext["missing_sheets"]:
                report.append(f"- {m}")
        report.append("")

    report.append("## 资质清单")
    missing_q = [r for r in quals["qualifications"] if not r["_resolved_path"]]
    report.append(f"- 条目数: {len(quals['qualifications'])}；原件缺失 {len(missing_q)} 条")
    for r in quals["qualifications"]:
        if r["_resolved_path"]:
            report.append(f"- ✅ {r.get('类别','')} | {r.get('证书完整名称','')} → {r['_resolved_path']}")
        else:
            report.append(f"- {MARK_MISSING} {r.get('类别','')} | {r.get('证书完整名称','')} | 台账路径: {r.get('文件路径（资料库内路径）','') or '(空)'} | {r['_resolve_note']}")
    report.append("")
    report.append("## 供应商入围资料清单")
    missing_m = [r for r in mats["materials"] if not r["_resolved_path"]]
    fixed = [r for r in mats["materials"] if r["_resolved_path"] and "递归定位" in (r.get("_resolve_note") or "")]
    report.append(f"- 条目数: {len(mats['materials'])}；原件缺失 {len(missing_m)} 条；路径修复(递归定位) {len(fixed)} 条")
    for r in missing_m:
        report.append(f"- {MARK_MISSING} {r.get('文件名','')} | 台账路径: {r.get('文件完整路径','') or '(空)'}")

    # 写文件
    outputs = [
        ("ledgers_company.json", company),
        ("ledgers_qualifications.json", quals),
        ("ledgers_materials.json", mats),
    ]
    if company_ext is not None:
        outputs.append(("ledgers_company_extended.json", company_ext))

    for name, data in outputs:
        with open(os.path.join(out, name), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out, "ledgers_verify_report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(report))

    print(f"✅ 台账解析完成。输出目录: {out}")
    print(f"   公司信息键值对 {len(company['company_info'])} | 资质 {len(quals['qualifications'])} (缺失 {len(missing_q)}) | 资料 {len(mats['materials'])} (缺失 {len(missing_m)}, 路径修复 {len(fixed)})")
    if company_ext is not None:
        n_sheets = len(company_ext["sheets"])
        n_missing = len(company_ext["missing_sheets"])
        print(f"   多工作表解析: {n_sheets} 个 sheet 匹配，{n_missing} 个未识别")
    if missing_q or missing_m:
        print(f"⚠️ 存在缺失项，详见 {os.path.join(out, 'ledgers_verify_report.md')}，请人工补充，禁止编造。")


if __name__ == "__main__":
    main()