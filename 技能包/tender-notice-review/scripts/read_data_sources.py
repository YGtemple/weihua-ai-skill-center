# -*- coding: utf-8 -*-
"""
只读读取两份权威数据源（CA登记表.xlsx / 重点平台清单.xlsx），输出 JSON。
硬性约束：本脚本严格只读，绝不写入、修改、移动、删除任何数据源文件。

用法：
  python read_data_sources.py [--check-date YYYY-MM-DD] [--out 输出.json]
    --check-date  指定校验基准日期（默认今天），用于判定CA证书有效期状态
    --out         结果同时写入该 JSON 文件（不指定则仅打印到 stdout）
"""
import argparse
import datetime as dt
import json
import re
import sys

import openpyxl

CA_XLSX = r"E:\卫华供应商入围资料库\CA登记表.xlsx"
PLATFORM_XLSX = r"E:\卫华供应商入围资料库\重点平台清单.xlsx"

# CA临近到期预警阈值（天）
EXPIRE_WARN_DAYS = 90


def parse_date(val):
    """把各种格式的日期解析为 date；无法解析返回 None。"""
    if val is None:
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val
    s = str(val).strip()
    if not s:
        return None
    # 去掉常见前缀如 "起始"、"截止" 等文字
    s = re.sub(r"[^\d.\-/年月日]", "", s)
    s = s.replace("年", ".").replace("月", ".").replace("日", "").replace("/", ".").replace("-", ".")
    parts = [p for p in s.split(".") if p]
    if len(parts) >= 3:
        try:
            return dt.date(int(parts[0]), int(parts[1]), int(parts[2]))
        except ValueError:
            return None
    return None


def fmt(d):
    return d.strftime("%Y-%m-%d") if d else None


def read_ca(check_date):
    wb = openpyxl.load_workbook(CA_XLSX, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 1:
            continue  # 表头：CA名称 | 起始时间 | 截止时间 | 备注
        name = str(row[0]).strip() if row[0] is not None and str(row[0]).strip() else None
        if not name:
            continue
        start = parse_date(row[1] if len(row) > 1 else None)
        end = parse_date(row[2] if len(row) > 2 else None)
        note = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
        item = {
            "row": i,
            "name": name,
            "start": fmt(start),
            "end": fmt(end),
            "note": note,
        }
        # 有效期状态判定（仅客观日期比对，不做任何主观推断）
        if check_date and end:
            days_left = (end - check_date).days
            if days_left < 0:
                item["status"] = "已过期"
                item["days_left"] = days_left
            elif days_left <= EXPIRE_WARN_DAYS:
                item["status"] = f"即将过期（剩余{days_left}天）"
                item["days_left"] = days_left
            else:
                item["status"] = "有效"
                item["days_left"] = days_left
            if start and check_date < start:
                item["status"] = "尚未生效"
        elif check_date and not end:
            item["status"] = "截止时间无法解析【需人工重点复核】"
        else:
            item["status"] = "未校验"
        rows.append(item)
    wb.close()
    return rows


def read_platforms():
    wb = openpyxl.load_workbook(PLATFORM_XLSX, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 1:
            continue  # 表头：网站注册名称 | 网址
        name = str(row[0]).strip() if row[0] is not None and str(row[0]).strip() else None
        if not name:
            continue
        url = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        rows.append({"row": i, "name": name, "url": url})
    wb.close()
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check-date", default=None, help="校验基准日期 YYYY-MM-DD，默认今天")
    ap.add_argument("--out", default=None, help="结果写入的JSON文件路径")
    args = ap.parse_args()

    check_date = None
    if args.check_date:
        try:
            check_date = dt.datetime.strptime(args.check_date, "%Y-%m-%d").date()
        except ValueError:
            print("错误：--check-date 格式应为 YYYY-MM-DD", file=sys.stderr)
            sys.exit(2)
    else:
        check_date = dt.date.today()

    result = {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "check_date": fmt(check_date),
        "sources": {"ca": CA_XLSX, "platform": PLATFORM_XLSX},
        "disclaimer": "AI仅辅助分析，所有内容必须人工逐条复核原件后才可使用，禁止直接对外提交。",
        "ca_count": 0,
        "platform_count": 0,
        "ca_list": [],
        "platform_list": [],
    }
    try:
        result["ca_list"] = read_ca(check_date)
        result["ca_count"] = len(result["ca_list"])
    except Exception as e:
        result["ca_error"] = f"读取CA登记表失败：{e}"
    try:
        result["platform_list"] = read_platforms()
        result["platform_count"] = len(result["platform_list"])
    except Exception as e:
        result["platform_error"] = f"读取重点平台清单失败：{e}"

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"已写入 {args.out}")
    print(text)


if __name__ == "__main__":
    main()
