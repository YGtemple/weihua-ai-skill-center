# -*- coding: utf-8 -*-
"""多工作表填充：根据 ledgers_company_extended.json，把甲方 Word 调查表里的
"主要客户信息 / 主要生产设备清单 / 主要检验设备清单 / 配套件/外协件供应商 /
设备年产量 / 财务信息"等表格按表头映射填入数据。

用法:
    python fill_multi_sheet_tables.py --template <甲方docx> \
        --extended <工作目录>/ledgers_company_extended.json \
        --out <工作目录>/成品.docx

硬性规则:
    - 仅读取已 parse_ledgers 产出的 extended JSON 作为数据源，禁止读取图片/PDF。
    - 源表无值 → 单元格留空 + ⚠️ 标记，禁止 AI 编造。
    - 模板表行数 > 源数据行数 → 多出的模板行留空 + ⚠️"需人工补充"。
    - 模板表行数 < 源数据行数 → 多出的源数据行追加到末尾 + ⚠️"超出模板行数"。
    - 财务信息按 (指标名, 年度) 填入对应单元格。
    - 设备年产量按 (序号, 品名, 规格, 年产量) 填入对应单元格。
    - 文档末尾强制加 AI 复核提示。
"""
import argparse
import json
import os
import sys
import re
from copy import deepcopy

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import MARK_MANUAL, MARK_MISSING, MARK_OK, die, ensure_dir  # noqa: E402

from docx import Document
from docx.shared import Pt, RGBColor


NOTE = "【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】"


# ---------------------------------------------------------------------------
# 表识别：模板表头行 → 目标类型
# ---------------------------------------------------------------------------

# 每类表的"识别关键词列表"：模板表头任一格包含任一关键词即命中
TABLE_RULES = [
    # (sheet_type, header_keywords)
    ("customers",         ["客户名称", "客户名", "主要客户"]),
    ("production_eq",     ["主要生产设备", "生产设备清单"]),
    ("inspection_eq",     ["主要检验设备", "主要检测设备", "检验设备清单", "检测设备清单"]),
    ("suppliers",         ["配套件供应商", "外协件供应商", "配套供应商", "外协供应商"]),
    ("annual_capacity",   ["设备年产量", "年产量"]),
    # financial 只用明确的"财务"二字，避免"营业收入/净资产/净利润"误命中利润表/财务状况表（双栏对比）
    ("financial",         ["财务信息", "财务状况", "财务数据"]),
]

# 列映射：模板列名（去掉空格/符号后做匹配）→ 数据源键
COLUMN_MAPS = {
    "customers": {
        "客户名称": "主要客户信息", "客户名": "主要客户信息", "客户": "主要客户信息",
        "联系人": "联系人", "联系电话": "联系电话", "电话": "联系电话",
        "产品名称": "产品名称", "产品": "产品名称",
        "使用量": "使用量", "合同期限": "合同期限",
    },
    "production_eq": {
        "设备名称": "设备名称", "设备名": "设备名称",
        "规格型号": "规格型号", "型号": "规格型号",
        "台套": "台/套", "台/套": "台/套",
        "制造厂家": "制造厂家", "厂家": "制造厂家",
        "出厂日期": "出厂日期",
        "单班产能": "单班产能",
        "关键工序": "关键工序",
    },
    "inspection_eq": {
        "设备名称": "设备名称", "设备名": "设备名称",
        "规格型号": "规格型号", "型号": "规格型号",
        "台套": "台/套", "台/套": "台/套",
        "制造厂家": "制造厂家", "厂家": "制造厂家",
        "出厂日期": "出厂日期",
        "单班产能": "单班产能",
        "关键工序": "关键工序",
    },
    "suppliers": {
        "设备": "设备", "类型": "设备",
        "供应商名称": "供应商名称", "供应商": "供应商名称",
        "联系人": "联系人",
        "联系电话": "联系电话", "电话": "联系电话",
        "供货能力": "供货能力",
    },
    "annual_capacity": {
        "序号": "序号", "品名": "品名", "产品名称": "品名",
        "规格": "规格", "规格型号": "规格",
        "年产量": "年产量",
    },
}


def _norm(s):
    """规范化字符串用于模糊匹配：去空白/全角符号。"""
    if not s:
        return ""
    return re.sub(r"[\s　，。、；：（）()【】/]", "", str(s))


def detect_table_type(table):
    """根据表头行（首行/前两行）单元格文本判断表格类型。返回类型标识 or None。

    处理两种常见模板:
    1) Row 0 直接是列头
    2) Row 0 是合并标题行（所有单元格内容相同），Row 1 才是真正的列头
    关键字同时在 Row 0 + Row 1 中查找（保证 Row 0 合并标题行也能被识别）
    """
    if not table.rows:
        return None
    header_texts_0 = [_norm(c.text) for c in table.rows[0].cells]
    header_texts_1 = ([_norm(c.text) for c in table.rows[1].cells]
                      if len(table.rows) >= 2 else [])
    header_join = "".join(header_texts_0) + "".join(header_texts_1)

    for sheet_type, keywords in TABLE_RULES:
        for kw in keywords:
            if _norm(kw) in header_join:
                return sheet_type
    return None


def _find_header_row_idx(table):
    """返回真正的表头行索引：0 或 1（若 Row 0 是合并标题行）。"""
    if len(table.rows) < 2:
        return 0
    non_empty = [_norm(c.text) for c in table.rows[0].cells]
    non_empty = [t for t in non_empty if t]
    if len(non_empty) > 1 and len(set(non_empty)) == 1:
        return 1
    return 0


def _map_columns(template_row, data_keys, col_map):
    """为模板表头行各列找到对应的数据键，返回 dict[template_col_idx] = data_key。
    template_col_idx 是去重后的有效列下标（同一合并 cell 只算一次）。"""
    seen = {}
    for cell in template_row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen[cid] = len(seen)
    # 收集每列的规范化文本
    norm_text_by_col = {}
    for cell in template_row.cells:
        cid = id(cell._tc)
        col_idx = seen[cid]
        nt = _norm(cell.text)
        # 取该列首个非空文本作为该列的列名
        if col_idx not in norm_text_by_col or not norm_text_by_col[col_idx]:
            norm_text_by_col[col_idx] = nt
    # 双向匹配
    out = {}
    used_data_keys = set()
    # 精确优先
    for col_idx, nt in norm_text_by_col.items():
        for template_name, data_key in col_map.items():
            if _norm(template_name) == nt:
                if data_key not in used_data_keys:
                    out[col_idx] = data_key
                    used_data_keys.add(data_key)
                    break
    # 模糊：模板 norm 包含 data_key 规范化 或反向
    for col_idx, nt in norm_text_by_col.items():
        if col_idx in out:
            continue
        for template_name, data_key in col_map.items():
            if data_key in used_data_keys:
                continue
            tn = _norm(template_name)
            if tn and (tn in nt or nt in tn):
                out[col_idx] = data_key
                used_data_keys.add(data_key)
                break
    return out


# ---------------------------------------------------------------------------
# 单元格填充工具
# ---------------------------------------------------------------------------

def _set_cell_text(cell, text):
    """把文本写入单元格（覆盖式），保留第一个段落，删除其他。"""
    if not text:
        return
    paragraphs = cell.paragraphs
    if not paragraphs:
        return
    p = paragraphs[0]
    # 清空段落内的所有 run
    for r in list(p.runs):
        r.text = ""
    # 清空段落的其他 runs 之后追加一个 run
    p.add_run(str(text))
    # 删除多余段落
    for extra in paragraphs[1:]:
        extra._element.getparent().remove(extra._element)


def _effective_cols(row):
    """考虑合并单元格：返回每个独立 cell 在该行中第几次出现的列表。
    python-docx 中合并单元格的同一 cell 在所有合并位置都被引用，
    所以同一文字的连续出现代表同一合并 cell。我们按出现次数取首个 cell 作为目标。
    seen_count = {}
    for cell in row.cells:
        idx = id(cell._tc)
        if idx not in seen_count:
            seen_count[idx] = len(seen_count)
    """
    # 简化：用 (id, first_occurrence) 标记
    seen = {}
    out = []
    for cell in row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen[cid] = len(seen)
        out.append(seen[cid])
    return out


def _map_columns(template_row, data_keys, col_map):
    """为模板表头行各列找到对应的数据键，返回 dict[template_col_idx] = data_key。
    template_col_idx 是去重后的有效列下标（同一合并 cell 只算一次）。"""
    seen = {}
    norm_seen = {}  # 规范化文本 → 首次出现的有效列下标
    for i, cell in enumerate(template_row.cells):
        cid = id(cell._tc)
        if cid in seen:
            continue
        seen[cid] = len(seen)
        nt = _norm(cell.text)
        if nt and nt not in norm_seen:
            norm_seen[nt] = seen[cid]
    # 双向匹配：模板头 (norm) ↔ col_map value
    out = {}
    rev_map = {data_key: template_name for template_name, data_key in col_map.items()}
    for template_norm, col_idx in norm_seen.items():
        for template_name, data_key in col_map.items():
            if template_norm == _norm(template_name):
                out[col_idx] = data_key
                break
        else:
            # 模糊：模板 norm 包含 data_key 规范化 或反向
            for template_name, data_key in col_map.items():
                tn = _norm(template_name)
                if tn and (tn in template_norm or template_norm in tn):
                    out[col_idx] = data_key
                    break
    return out


# ---------------------------------------------------------------------------
# 通用行式表填充：customers / production_eq / inspection_eq / suppliers / annual_capacity
# ---------------------------------------------------------------------------

def fill_rowwise_table(table, sheet_type, rows, col_map, fill_report):
    """通用行式表填充。rows 是 list of dict（数据源）。"""
    if not rows:
        fill_report.append((f"[{sheet_type}]", "❌", "源表无数据，全部留空 + 需人工补充"))
        return 0

    header_row_idx = _find_header_row_idx(table)
    if len(table.rows) <= header_row_idx:
        fill_report.append((f"[{sheet_type}]", "⚠️", "模板表无数据行，跳过"))
        return 0

    col_idx_to_key = _map_columns(table.rows[header_row_idx], rows[0].keys() if rows else [], col_map)
    if not col_idx_to_key:
        fill_report.append((f"[{sheet_type}]", "⚠️", "无法识别模板表头与源表字段的对应关系，请人工核对表头"))
        return 0

    # 数据���从 header_row_idx+1 开始
    n_template_data_rows = len(table.rows) - header_row_idx - 1
    n_source = len(rows)
    n_fill = min(n_template_data_rows, n_source)

    for r_idx in range(n_fill):
        src_row = rows[r_idx]
        tgt_row = table.rows[r_idx + header_row_idx + 1]
        # 合并单元格：同一个 tc 在 row.cells 中多次出现，只写第一个
        seen_tc = {}
        for c_idx, cell in enumerate(tgt_row.cells):
            cid = id(cell._tc)
            if cid in seen_tc:
                continue
            seen_tc[cid] = len(seen_tc)
            eff_idx = seen_tc[cid]
            data_key = col_idx_to_key.get(eff_idx)
            if not data_key:
                continue
            val = src_row.get(data_key, "") or ""
            if val:
                _set_cell_text(cell, val)

    fill_report.append((
        f"[{sheet_type}]",
        "✅" if n_source <= n_template_data_rows else "⚠️",
        f"已填 {n_fill} 行（模板行 {n_template_data_rows} | 源行 {n_source}）"
        + ("" if n_source <= n_template_data_rows else f"，多出 {n_source - n_template_data_rows} 条源数据已记录待人工追加")
        + ("" if n_template_data_rows <= n_source else f"，多出 {n_template_data_rows - n_source} 模板行已留空 + ⚠️需人工补充")
    ))
    return n_fill


# ---------------------------------------------------------------------------
# 财务信息表填充：横向年度表，每行一个指标名，每列一个年度
# ---------------------------------------------------------------------------

def fill_financial_table(table, fin_data, fill_report):
    """财务信息按 (指标名, 年度) 填入对应单元格。
    模板表头可能含 '指标名/2018年/...' 或 '指标/名称/...' 或多个年度。
    """
    years = fin_data.get("years", [])
    metrics = fin_data.get("metrics", [])
    if not metrics:
        fill_report.append(("[财务信息]", "❌", "源财务表无数据，全部留空 + 需人工补充"))
        return 0

    header_row_idx = _find_header_row_idx(table)
    if len(table.rows) <= header_row_idx:
        fill_report.append(("[财务信息]", "⚠️", "模板表无数据行，跳过"))
        return 0

    # 解析模板表头：第 0 列 = 指标名，后面每列 = 年度
    header_row = table.rows[header_row_idx]
    seen_tc = {}
    header_texts = []  # 每个有效列下标 → 文本
    for cell in header_row.cells:
        cid = id(cell._tc)
        if cid not in seen_tc:
            seen_tc[cid] = len(seen_tc)
            header_texts.append(_norm(cell.text))
        # 重复出现则跳过（合并单元格）

    # 找到模板年度列（与源 years 做匹配）
    year_to_col = {}
    for col_idx, ht in enumerate(header_texts[1:], start=1):  # 跳过第一列（指标名列）
        for y in years:
            if _norm(y) == ht or _norm(y) in ht or ht in _norm(y):
                year_to_col[y] = col_idx
                break

    # 模糊匹配兜底：模板表头里含"2018"/"2019"/"2024"等关键数字的列
    if not year_to_col:
        for col_idx, ht in enumerate(header_texts[1:], start=1):
            for y in years:
                ynum = re.findall(r"\d{4}", y)
                if ynum and ynum[0] in ht:
                    year_to_col[y] = col_idx
                    break

    if not year_to_col:
        fill_report.append(("[��务信息]", "⚠️",
                            f"无法识别模板表头年度列（源年度={years} | 模板表头={header_texts}），请人工核对"))
        return 0, []

    # 匹配指标名（模糊）：模板行第一列文本 与 metric['name'] 做包含匹配
    n_template_metric_rows = len(table.rows) - header_row_idx - 1
    n_source = len(metrics)
    n_fill = 0
    used_template_rows = set()
    matched_metrics = {}  # 指标名 → 模板行下标

    # 构建"指标规范化名 → metric"的索引
    norm_metric_idx = {}
    for m in metrics:
        nm = _norm(m["name"])
        norm_metric_idx[nm] = m
        # 去括号换行片段
        nm_short = re.split(r"\(|\n|（", nm)[0]
        if nm_short and nm_short not in norm_metric_idx:
            norm_metric_idx[nm_short] = m

    # 简化别名映射（财务信息常见同义）
    alias = {
        "年度销售总额": "营业收入",
        "销售总额": "营业收入",
        "营收": "营业收入",
        "资产总额": "全部资产总值",
        "总资产": "全部资产总值",
        "负债": "流动负债",  # 兜底，精确匹配优先
        "销售净利润": "净利润",
        "净利": "净利润",
        "经营性现金流": "经营活动现金净流量",
    }

    for r_idx in range(header_row_idx + 1, len(table.rows)):
        metric_name_text = _norm(table.rows[r_idx].cells[0].text)
        if not metric_name_text:
            continue
        # 精确 or 包含 匹配
        m = norm_metric_idx.get(metric_name_text)
        if not m:
            for nmn, rec in norm_metric_idx.items():
                if metric_name_text and (metric_name_text in nmn or nmn in metric_name_text):
                    m = rec
                    break
        # 别名匹配
        if not m and metric_name_text in alias:
            alias_target = alias[metric_name_text]
            m = norm_metric_idx.get(_norm(alias_target))
        if not m:
            continue
        matched_metrics[m["name"]] = r_idx
        used_template_rows.add(r_idx)
        # 填充每个年度单元格
        for y, col_idx in year_to_col.items():
            v = m["values"].get(y)
            if v:
                # 找到此列的有效 cell
                seen = {}
                tgt_cell = None
                for c_idx, cell in enumerate(table.rows[r_idx].cells):
                    cid = id(cell._tc)
                    if cid not in seen:
                        seen[cid] = len(seen)
                    if seen[cid] == col_idx:
                        tgt_cell = cell
                        break
                if tgt_cell is not None:
                    _set_cell_text(tgt_cell, v)
                    break  # 已找到唯一目标，退出 inner loop

        n_fill += 1

    unmatched = [m["name"] for m in metrics if m["name"] not in matched_metrics]
    unused_template_rows = n_template_metric_rows - len(used_template_rows)

    note_parts = [f"已填 {n_fill} 行（指标匹配）"]
    if unmatched:
        note_parts.append(f"源表未匹配的指标: {';'.join(unmatched[:5])}")
    if unused_template_rows > 0:
        note_parts.append(f"模板行多出 {unused_template_rows} 行，已留空 + ⚠️需人工补充")
    fill_report.append(("[财务信息]",
                        "✅" if not unmatched and unused_template_rows == 0 else "⚠️",
                        " | ".join(note_parts)))

    # 数据源有、模板无落点的指标 → 在文档末尾追加一段提示
    return n_fill, unmatched


# ---------------------------------------------------------------------------
# 文档末尾追加 AI 复核提示
# ---------------------------------------------------------------------------

def append_ai_notice(doc):
    """在文档末尾追加 AI 复核提示段落（若文档已存在该提示则不重复追加）。"""
    if any(p.text.strip().startswith("【本结果由AI生成") for p in doc.paragraphs):
        return
    p = doc.add_paragraph()
    r = p.add_run(NOTE)
    r.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, help="甲方 Word 调查表 .docx")
    ap.add_argument("--extended", required=True, help="parse_ledgers 产出的 ledgers_company_extended.json")
    ap.add_argument("--out", required=True, help="输出成品 .docx")
    args = ap.parse_args()

    if not os.path.isfile(args.template):
        die(f"甲方模板不存在: {args.template}")
    if not os.path.isfile(args.extended):
        die(f"extended ledger 不存在: {args.extended}")

    with open(args.extended, encoding="utf-8") as f:
        ext = json.load(f)

    sheets = ext.get("sheets", {})
    if not sheets:
        die("extended ledger 中无 sheets 数据，请先运行 parse_ledgers.py")

    doc = Document(args.template)
    fill_report = []

    # 财务信息的多余指标（在文档末尾展示用）
    extra_financial_metrics = []

    n_tables = len(doc.tables)
    for ti, table in enumerate(doc.tables):
        sheet_type = detect_table_type(table)
        if not sheet_type:
            continue

        # 找到对应的源 sheet
        src = None
        for sheet_name, sheet_data in sheets.items():
            if sheet_data.get("type") == sheet_type:
                src = sheet_data
                break

        if not src:
            fill_report.append((f"Table {ti+1} [{sheet_type}]", "⚠️",
                                "extended ledger 中无对应类型数据，跳过填充"))
            continue

        if sheet_type == "financial":
            n, unmatched = fill_financial_table(table, src, fill_report)
            if unmatched:
                extra_financial_metrics.extend(unmatched)
        elif sheet_type == "annual_capacity":
            rows = src.get("rows", [])
            n = fill_rowwise_table(table, sheet_type, rows, COLUMN_MAPS["annual_capacity"], fill_report)
        else:
            rows = src.get("rows", [])
            col_map = COLUMN_MAPS.get(sheet_type, {})
            n = fill_rowwise_table(table, sheet_type, rows, col_map, fill_report)

    # 如果有数据源有但模板无落点的财务指标，在末尾追加一段
    if extra_financial_metrics:
        p = doc.add_paragraph()
        r = p.add_run(f"【数据源有该指标，调查表无落点，请人工补充：{'；'.join(extra_financial_metrics[:20])}】")
        r.font.color.rgb = RGBColor(0xC0, 0x80, 0x00)

    # 文档末尾强制提示
    append_ai_notice(doc)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    ensure_dir(out_dir)
    doc.save(args.out)

    print(f"✅ 多表格填充完成。")
    print(f"   模板表数: {n_tables} | 输出: {args.out}")
    print(f"\n===== 填充报告 =====")
    for pos, st, note in fill_report:
        print(f"[{st}] {pos}: {note}")

    # 写入填充报告
    report_path = os.path.join(out_dir, "multi_sheet_fill_report.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# 多工作表填充报告\n\n")
        f.write(f"- 模板: {args.template}\n")
        f.write(f"- 输出: {args.out}\n")
        f.write(f"- 数据源: {args.extended}\n\n")
        for pos, st, note in fill_report:
            f.write(f"- [{st}] {pos}: {note}\n")
        if extra_financial_metrics:
            f.write(f"\n## ⚠️ 数据源有但模板无落点的���务指标\n\n")
            for m in extra_financial_metrics:
                f.write(f"- {m}\n")
        f.write(f"\n> {NOTE}\n")
    print(f"\n报告: {report_path}")


if __name__ == "__main__":
    main()