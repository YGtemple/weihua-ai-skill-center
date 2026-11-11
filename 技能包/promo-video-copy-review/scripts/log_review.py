#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""审稿留痕与闭环：审完一篇记一行，驳回理由自动归并进校准文件，并输出累计统计。

用法（最小）：
  python log_review.py --id S-002 --title "xx宣传视频" --score 54 --verdict 驳回重写 --p0 4

完整示例：
  python log_review.py --id S-002 --title "xx宣传视频" --type 企业动态 --carrier 27s竖屏视频 \
      --chars 199 --duration 27.4 --score 54 --verdict 驳回重写 --p0 4 --p1 3 \
      --minutes 6 --cross-check --reason "字幕与文本名单不一致" --dimension B2 \
      --core "字数超载62%；名单不一致" --pending "名单职务、主办方全称" --submitter 张三

设计原则：
- 台账是数据，放在 ledger/ 下，与技能标准文件分开；首次运行自动从 assets 模板复制表头。
- 驳回理由写入 references/calibration-log.md 的「四、驳回理由库」表格，占位的"（待积累）"会被替换。
- 不臆造耗时：--minutes 由审核人填实际用时，没填就不参与平均耗时统计。

架构说明（标准与代码分离）：
- 台账表头 HEADER、附加列 EXTRA、维度归类 DIM_TO_TYPE 在 config/ledger-schema.json，加列只改那里。
- 结论选项 --verdict 从 config/scoring.json 的 verdicts 读取，不与评分结论线两处维护。
- 配置缺失或损坏时给出清晰错误并退出。
"""

import argparse
import csv
import datetime
import json
import os
import re
import sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(SKILL_DIR, "config")
DEFAULT_LEDGER = os.path.join(SKILL_DIR, "ledger", "审核台账.csv")
TEMPLATE = os.path.join(SKILL_DIR, "assets", "审核台账模板.csv")
CALIBRATION = os.path.join(SKILL_DIR, "references", "calibration-log.md")


def load_config(filename):
    """读 config/ 下的 JSON；缺失/损坏都给出清晰错误并退出。"""
    path = os.path.join(CONFIG_DIR, filename)
    if not os.path.exists(path):
        sys.exit(f"缺少配置文件：{path}\n请确认 {filename} 存在于 config/ 目录后再运行。")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        sys.exit(f"配置文件不是合法 JSON：{path}\n详情：{e}")


SCHEMA = load_config("ledger-schema.json")
SCORING = load_config("scoring.json")

# 台账表头与附加列，外置到 config/ledger-schema.json
HEADER = SCHEMA["header"]
EXTRA = SCHEMA.get("extra_columns", [])
DIM_TO_TYPE = SCHEMA.get("dim_to_type", {})
DIM_FALLBACK = SCHEMA.get("dim_fallback", "效果")

# 结论选项复用 config/scoring.json，避免两处维护
VERDICT_CHOICES = [v["label"] for v in SCORING.get("verdicts", [])]


def ensure_ledger(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        with open(TEMPLATE, "r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        header = rows[0] + EXTRA
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerow(header)
        return header
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        sys.exit("台账文件为空，请检查：" + path)
    header = rows[0]
    if len(header) == len(HEADER):
        # 旧版 21 列表头，自动补附加列
        header = header + EXTRA
        new_rows = [header] + [r + [""] * len(EXTRA) for r in rows[1:]]
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(new_rows)
    return header


def read_rows(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))


def append_calibration(reason, dimension, date, gid):
    """把驳回理由写进 calibration-log.md 的驳回理由库。"""
    if not os.path.exists(CALIBRATION):
        return "未找到 calibration-log.md，理由未归档"
    with open(CALIBRATION, "r", encoding="utf-8") as f:
        text = f.read()

    m = re.search(r"(## 四、驳回理由库[\s\S]*?)(\n## |\Z)", text)
    if not m:
        return "calibration-log.md 中未找到「四、驳回理由库」小节，理由未归档"

    section = m.group(1)
    dtype = DIM_TO_TYPE.get((dimension or "")[:1].upper(), DIM_FALLBACK)
    dim = dimension or "待归类"
    cell = f"{reason}（{gid}，{date}）"

    lines = section.splitlines()
    replaced = False
    for i, line in enumerate(lines):
        if line.startswith("|") and "（待积累）" in line:
            cols = [c.strip() for c in line.strip().strip("|").split("|")]
            if cols and cols[0] == dtype:
                lines[i] = f"| {dtype} | {cell} | {dim} |"
                replaced = True
                break
    if not replaced:
        # 插到本小节最后一个表格行之后
        last_tbl = -1
        for i, line in enumerate(lines):
            if line.startswith("|"):
                last_tbl = i
        new_line = f"| {dtype} | {cell} | {dim} |"
        if last_tbl >= 0:
            lines.insert(last_tbl + 1, new_line)
        else:
            lines.append(new_line)

    new_section = "\n".join(lines)
    text = text[:m.start(1)] + new_section + text[m.end(1):]
    with open(CALIBRATION, "w", encoding="utf-8") as f:
        f.write(text)
    return f"理由已归档：{dtype} / {dim}"


def stats(path):
    rows = read_rows(path)
    if len(rows) < 2:
        return "暂无数据"
    header = rows[0]
    idx = {name: i for i, name in enumerate(header)}
    data = [r for r in rows[1:] if r and r[0].strip()]

    def get(r, name):
        i = idx.get(name)
        return r[i].strip() if i is not None and i < len(r) else ""

    total = len(data)
    scores = [float(get(r, "总分")) for r in data if get(r, "总分").replace(".", "", 1).isdigit()]
    p0 = sum(int(get(r, "必改P0数")) for r in data if get(r, "必改P0数").isdigit())
    minutes = [float(get(r, "审核耗时分钟")) for r in data
               if get(r, "审核耗时分钟").replace(".", "", 1).isdigit()]
    checked = sum(1 for r in data if get(r, "是否核片") == "是")
    verdicts = {}
    for r in data:
        v = get(r, "结论") or "未填"
        verdicts[v] = verdicts.get(v, 0) + 1

    out = [
        f"累计审稿 {total} 篇",
        f"平均得分 {sum(scores) / len(scores):.1f}" if scores else "平均得分 暂无",
        f"拦下 P0 共 {p0} 项",
        f"核片率 {checked}/{total}（{checked / total * 100:.0f}%）",
        f"平均审核耗时 {sum(minutes) / len(minutes):.1f} 分钟（{len(minutes)} 篇有记录）"
        if minutes else "平均审核耗时 暂无（用 --minutes 记录实际用时）",
        "结论分布 " + "、".join(f"{k} {v}" for k, v in sorted(verdicts.items())),
    ]
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description="审稿留痕与闭环")
    ap.add_argument("--id")
    ap.add_argument("--title")
    ap.add_argument("--score", type=float)
    ap.add_argument("--verdict", choices=VERDICT_CHOICES)
    ap.add_argument("--type", default="企业动态")
    ap.add_argument("--carrier", default="")
    ap.add_argument("--chars", default="")
    ap.add_argument("--duration", default="")
    ap.add_argument("--submitter", default="")
    ap.add_argument("--p0", type=int, default=0)
    ap.add_argument("--p1", type=int, default=0)
    ap.add_argument("--core", default="", help="核心问题，一句话")
    ap.add_argument("--pending", default="", help="待补信息")
    ap.add_argument("--minutes", type=float, default=None, help="实际审核耗时（分钟），不填则不统计")
    ap.add_argument("--cross-check", action="store_true", help="本次是否核对了成片")
    ap.add_argument("--hits", default="", help="机检风险词命中数")
    ap.add_argument("--reason", default="", help="驳回理由（给了就自动归档进校准文件）")
    ap.add_argument("--dimension", default="", help="对应扣分项编号，如 B2")
    ap.add_argument("--date", default=datetime.date.today().isoformat())
    ap.add_argument("--ledger", default=DEFAULT_LEDGER)
    ap.add_argument("--stats-only", action="store_true", help="只看累计统计，不记新行")
    args = ap.parse_args()

    header = ensure_ledger(args.ledger)

    if args.stats_only:
        print(stats(args.ledger))
        return

    missing = [name for name, val in
               [("--id", args.id), ("--title", args.title),
                ("--score", args.score), ("--verdict", args.verdict)] if val is None]
    if missing:
        ap.error("记账时以下参数必填：" + "、".join(missing))

    row_map = {
        "编号": args.id, "收到日期": args.date, "提交人": args.submitter,
        "稿件标题": args.title, "稿件类型": args.type, "载体": args.carrier,
        "字数": str(args.chars), "成片时长秒": str(args.duration),
        "总分": str(args.score), "结论": args.verdict,
        "必改P0数": str(args.p0), "建议P1数": str(args.p1),
        "核心问题": args.core, "待补信息": args.pending,
        "处理时限": "", "状态": "已审", "复审结果": "未复审",
        "审核耗时分钟": "" if args.minutes is None else str(args.minutes),
        "是否核片": "是" if args.cross_check else "否",
        "机检命中数": str(args.hits),
    }
    row = [row_map.get(col, "") for col in header]

    with open(args.ledger, "a", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerow(row)
    print(f"已记账：{args.id}｜{args.title}｜{args.score} 分｜{args.verdict}｜P0 {args.p0}")

    if args.reason:
        print(append_calibration(args.reason, args.dimension, args.date, args.id))

    print("\n--- 累计统计 ---")
    print(stats(args.ledger))


if __name__ == "__main__":
    main()
