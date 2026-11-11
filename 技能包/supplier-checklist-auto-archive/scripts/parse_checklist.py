# -*- coding: utf-8 -*-
"""通用版-供应商资料清单自动归档：Step 1 自适应解析甲方资料要求/审核项 Excel。

功能：
  1. 读取上传 Excel 的所有工作表，选择数据量最大的表；
  2. 自动补全合并单元格内容到区域内的每一行；
  3. 自动识别【分类列】（合并单元格/重复层级列）与【要求列】（具体资料要求/审核项）；
  4. 原表无序号列时自动生成层级序号（1、1.1、1.2、2、2.1……）；
  5. 输出解析预览表（xlsx + json + 控制台 markdown），等待用户确认。

识别不到【要求列】→ 告警退出（exit 2）；识别不到【分类列】→ 降级为“未分类”并在预览中标注警告。
"""
import argparse
import json
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

CATEGORY_HEADER_HINTS = ["分类", "类别", "分类项", "大项", "项目分类", "一级"]
REQ_HEADER_HINTS = ["要求", "审核", "评价", "资料", "内容", "指标", "检查", "评分", "标准", "说明"]
SEQ_HEADER_HINTS = ["序号", "编号", "项号", "no", "No", "NO", "序"]
SEQ_VALUE_RE = re.compile(r"^\d{1,3}(\.\d{1,3})?$")


def norm(s):
    if s is None:
        return ""
    return re.sub(r"\s+", "", str(s)).strip()


def cell_text(v):
    if v is None:
        return ""
    return str(v).strip()


def pick_sheet(wb, sheet_name=None):
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise SystemExit(f"[告警] 找不到工作表 {sheet_name}")
        return wb[sheet_name]
    # 选非空单元格最多的表
    best, best_n = None, -1
    for ws in wb.worksheets:
        n = sum(1 for row in ws.iter_rows(values_only=True) for c in row if cell_text(c))
        if n > best_n:
            best, best_n = ws, n
    if best is None:
        raise SystemExit("[告警] Excel 中没有任何数据")
    return best


def build_grid(ws):
    """读取为二维文本网格，并补全合并单元格。"""
    max_c = ws.max_column or 1
    grid = []
    for row in ws.iter_rows(values_only=True):
        grid.append([cell_text(v) for v in row] + [""] * (max_c - len(row)))
    merged_cols = set()
    for rng in ws.merged_cells.ranges:
        tl = ws.cell(rng.min_row, rng.min_col).value
        val = cell_text(tl)
        merged_cols.add(rng.min_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if 1 <= r <= len(grid) and 1 <= c <= max_c:
                    grid[r - 1][c - 1] = val
    return grid, merged_cols, max_c


def find_header_row(grid):
    """探测表头行：只在前 3 行内寻找（表头不可能在表格中部，防止分类文字含
    「指标/评价」等关键词的行被误判为表头导致上方数据被丢弃）。
    条件：本行多为短文本，且（含表头关键词 或 下一行非空单元格更多/内容更长）。"""
    for i, row in enumerate(grid[:3]):
        vals = [v for v in row if v]
        if not vals or i + 1 >= len(grid):
            continue
        nxt = [v for v in grid[i + 1] if v]
        headerish = all(len(v) <= 12 for v in vals)
        has_hint = any(any(h in v for h in CATEGORY_HEADER_HINTS + REQ_HEADER_HINTS + SEQ_HEADER_HINTS) for v in vals)
        grows_below = len(nxt) > len(vals) or (sum(len(v) for v in nxt) > sum(len(v) for v in vals) * 1.5)
        if headerish and (has_hint or grows_below):
            return i
    return None  # 无表头


def column_stats(grid, header_idx, data_start):
    stats = []
    n_rows = len(grid) - data_start
    for c in range(len(grid[0]) if grid else 0):
        vals = [grid[r][c] for r in range(data_start, len(grid))]
        nonempty = [v for v in vals if v]
        distinct = []
        for v in nonempty:
            if v not in distinct:
                distinct.append(v)
        header = grid[header_idx][c] if header_idx is not None and header_idx < len(grid) else ""
        stats.append({
            "col": c,
            "header": header,
            "nonempty": len(nonempty),
            "distinct": len(distinct),
            "avg_len": (sum(len(v) for v in nonempty) / len(nonempty)) if nonempty else 0,
        })
    return stats


def detect_columns(grid, merged_cols, data_start, n_data_rows):
    """返回 (seq_cols, cat_col, req_col, warnings)"""
    warnings = []
    header_idx = find_header_row(grid)
    if header_idx is not None:
        data_start = header_idx + 1
    stats = column_stats(grid, header_idx, data_start)
    usable = [s for s in stats if s["nonempty"] >= max(2, n_data_rows * 0.2)]
    if not usable:
        raise SystemExit("[告警] 未识别到任何有效数据列，解析停止。请检查表格内容。")

    # 1) 序号列：表头含序号提示，或值几乎全为 1 / 1.1 型短编号
    seq_cols = set()
    for s in usable:
        if any(h in s["header"] for h in SEQ_HEADER_HINTS):
            seq_cols.add(s["col"])
        else:
            vals = [grid[r][s["col"]] for r in range(data_start, len(grid))]
            vals = [v for v in vals if v]
            if vals and len(vals) >= 3 and sum(1 for v in vals if SEQ_VALUE_RE.match(v)) / len(vals) > 0.8 and s["avg_len"] <= 6:
                seq_cols.add(s["col"])

    # 2) 分类列：参与合并单元格，或取值重复率高（distinct/nonempty 低）且非序号列
    cat_candidates = []
    for s in usable:
        if s["col"] in seq_cols:
            continue
        ratio = s["distinct"] / s["nonempty"] if s["nonempty"] else 1
        if (s["col"] + 1) in merged_cols or ratio < 0.6:
            cat_candidates.append((s["col"], ratio, s))
    # 表头提示加权
    cat_candidates.sort(key=lambda t: (not any(h in t[2]["header"] for h in CATEGORY_HEADER_HINTS), t[0]))
    cat_col = cat_candidates[0][0] if cat_candidates else None
    if cat_col is None:
        warnings.append("未识别到【分类列】（无合并单元格、无重复层级列），已降级为单一分类“未分类”，请人工确认预览结果。")

    # 3) 要求列：剩余列中文本量最大者；表头提示加权
    req_candidates = [s for s in usable if s["col"] not in seq_cols and s["col"] != cat_col]
    if not req_candidates:
        raise SystemExit("[告警] 未识别到【要求列】，无法进行文件匹配，解析停止。请确认表格包含资料要求/审核项内容。")
    req_candidates.sort(key=lambda s: (not any(h in s["header"] for h in REQ_HEADER_HINTS), -(s["avg_len"] * s["nonempty"])))
    req_col = req_candidates[0]["col"]
    if req_candidates[0]["avg_len"] < 3:
        raise SystemExit("[告警] 要求列内容过短或为空，解析结果不可靠，停止执行。")

    return header_idx, data_start, seq_cols, cat_col, req_col, warnings


def sanitize_name(s):
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip() or "未命名"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="甲方资料要求 Excel 路径")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--sheet", default=None, help="指定工作表名（默认自动选择数据量最大的表）")
    args = ap.parse_args()

    if not os.path.isfile(args.input):
        raise SystemExit(f"[告警] 输入文件不存在：{args.input}")
    os.makedirs(args.out, exist_ok=True)

    wb = openpyxl.load_workbook(args.input, data_only=True)
    ws = pick_sheet(wb, args.sheet)
    grid, merged_cols, max_c = build_grid(ws)

    # 去掉全空行/全空列
    grid = [row for row in grid if any(row)]
    if not grid:
        raise SystemExit("[告警] 表格为空")
    keep_cols = [c for c in range(max_c) if any(grid[r][c] for r in range(len(grid)))]
    grid = [[row[c] for c in keep_cols] for row in grid]
    remap = {old: new for new, old in enumerate(keep_cols)}
    merged_cols = {remap[c] for c in merged_cols if c in remap}

    n_data_rows_guess = max(1, len(grid) - 1)
    header_idx, data_start, seq_cols, cat_col, req_col, warnings = detect_columns(
        grid, merged_cols, 1, n_data_rows_guess)

    # 分类列前向补全（合并单元格已补，此处兜底处理“写一次不合并”的情况）
    if cat_col is not None:
        last = ""
        for r in range(data_start, len(grid)):
            if grid[r][cat_col]:
                last = grid[r][cat_col]
            else:
                grid[r][cat_col] = last

    # 生成层级序号 + 条目（同分类下与上一行完全相同的要求 = 原表合并单元格产物，自动折叠去重）
    items, cur_cat, major, minor, deduped = [], None, 0, 0, 0
    for r in range(data_start, len(grid)):
        cat = grid[r][cat_col] if cat_col is not None else ""
        req = grid[r][req_col]
        if not req:
            continue  # 无要求内容的行（纯分类标题行等）跳过
        if items and items[-1]["category_raw"] == (cat or "未分类") and items[-1]["requirement"] == req:
            deduped += 1
            continue
        if cat != cur_cat:
            major += 1
            minor = 0
            cur_cat = cat
        minor += 1
        seq = f"{major}.{minor}" if cat_col is not None else f"{major}"
        # 附加列（除序号/分类/要求外的其他列）原样保留
        extra = {f"列{get_column_letter(c + 1)}": grid[r][c]
                 for c in range(len(grid[r]))
                 if c not in seq_cols and c != cat_col and c != req_col and grid[r][c]}
        items.append({
            "seq": seq,
            "category": sanitize_name(cat) if cat else "未分类",
            "category_raw": (cat or "未分类"),
            "requirement": req,
            "extra": extra,
        })
    if deduped:
        warnings.append(f"已自动折叠原表合并单元格产生的连续重复行 {deduped} 处（同一分类下内容完全相同的要求合并为一条编号）。")

    if not items:
        raise SystemExit("[告警] 未解析出任何资料要求条目，停止执行。")

    # 输出 JSON（供 Step 2 匹配归档使用）
    result = {
        "source": os.path.abspath(args.input),
        "sheet": ws.title,
        "header_row": (header_idx + 1) if header_idx is not None else None,
        "category_col": (get_column_letter(cat_col + 1)) if cat_col is not None else None,
        "requirement_col": get_column_letter(req_col + 1),
        "warnings": warnings,
        "items": items,
    }
    json_path = os.path.join(args.out, "parsed_items.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 输出解析预览 xlsx
    preview_path = os.path.join(args.out, "解析预览.xlsx")
    pwb = openpyxl.Workbook()
    pws = pwb.active
    pws.title = "解析预览"
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    pws.append(["自动生成序号", "分类名称（已补全）", "要求内容"] + [k for k in (items[0]["extra"] or {})])
    for cell in pws[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
        cell.alignment = Alignment(vertical="center")
    for it in items:
        pws.append([it["seq"], it["category"], it["requirement"]] + list(it["extra"].values()))
    if warnings:
        r = pws.max_row + 2
        pws.cell(r, 1, "⚠️ " + "；".join(warnings)).fill = warn_fill
        pws.cell(r + 1, 1, "⚠️ 本解析结果由AI生成，请人工确认无误后，再执行文件匹配与归集。").fill = warn_fill
    for col, w in zip("ABC", (14, 30, 60)):
        pws.column_dimensions[col].width = w
    pwb.save(preview_path)

    # 控制台 markdown 预览
    print(f"[解析完成] 工作表「{ws.title}」，共 {len(items)} 条要求，输出：{preview_path}")
    print(f"  分类列：{result['category_col'] or '（未识别，降级为未分类）'} | 要求列：{result['requirement_col']}")
    for w in warnings:
        print("  [警告]", w)
    print()
    print("| 序号 | 分类 | 要求内容 |")
    print("|---|---|---|")
    for it in items:
        print(f"| {it['seq']} | {it['category']} | {it['requirement'][:50]} |")
    print()
    print("请人工确认以上解析预览无误后，再执行 Step 2（match_and_archive.py）进行文件匹配与归集。")


if __name__ == "__main__":
    main()
