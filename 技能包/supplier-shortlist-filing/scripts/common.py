# -*- coding: utf-8 -*-
"""供应商入围资料库 skill 公共工具：路径容错解析、日期序列号转换、统一标记。"""
import os
import sys
import json
import datetime

# 与用户约定一致的标记
MARK_OK = "✅已匹配"
MARK_MANUAL = "⚠️需人工补充"
MARK_MISSING = "❌资料库缺失"

DEFAULT_BASE = r"E:\卫华供应商入围资料库"
LEDGER_FILES = {
    "company": "供方能力调查基础信息表.xlsx",
    "qualification": "资质清单.xlsx",
    "material": "供应商入围资料清单.xlsx",
}


def die(msg):
    """读取/解析失败时明确告警退出，禁止编造。"""
    print("【告警】" + msg, file=sys.stderr)
    sys.exit(2)


def normalize_path(p):
    """把台账里可能的错误路径规范为反斜杠 Windows 路径字符串。"""
    if not p:
        return ""
    return str(p).replace("/", "\\").strip()


def resolve_file(base, ledger_path, filename):
    """优先用台账路径；路径无效或文件不存在时，按文件名在 base 下递归搜索。

    返回 (绝对路径 or None, 说明字符串)
    """
    ledger_path = normalize_path(ledger_path)
    # 1) 台账给出的绝对/相对路径
    if ledger_path:
        cand = ledger_path if os.path.isabs(ledger_path) else os.path.join(base, ledger_path)
        if os.path.isfile(cand):
            return os.path.abspath(cand), f"台账路径有效: {ledger_path}"
    # 2) 文件名递归搜索（容错台账路径缺陷，如缺分隔符的拼接路径）
    if filename:
        for root, _dirs, files in os.walk(base):
            for f in files:
                if f == filename:
                    full = os.path.abspath(os.path.join(root, f))
                    note = f"台账路径无效，按文件名递归定位: {full}"
                    return full, note
    return None, f"未找到文件: 台账路径={ledger_path or '(空)'}, 文件名={filename or '(空)'}"


def excel_serial_to_date(v):
    """Excel 日期序列号 → 'YYYY-MM-DD'；无法转换返回 None。"""
    if isinstance(v, (int, float)) and v > 1:
        try:
            d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))
            return d.strftime("%Y-%m-%d")
        except Exception:
            return None
    return None


def _is_date_format(fmt):
    """根据单元格 number_format 判断是否为日期类型。
    openpyxl 把 'yyyy-mm-dd' / 'yyyy年m月d日' / '[$-409]m/d/yyyy' 等识别为日期格式。
    """
    if not fmt:
        return False
    fmt_l = fmt.lower().replace("[$-409]", "")
    return any(tok in fmt_l for tok in ("yy", "mm", "dd", "m月", "d日"))


def cell_str(v, number_format=None):
    """单元格值 → 字符串。
    - 如果 number_format 明确是日期型 → 转 YYYY-MM-DD
    - 否则保持数值（不去尾零、保留原数字格式），不要把使用量 9656 误转成 1926-06-08
    - 字符串原样返回
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, (int, float)):
        if _is_date_format(number_format):
            d = excel_serial_to_date(v)
            if d:
                return d
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip()


def cell_str_legacy(v):
    """兼容旧逻辑：无 number_format 时也自动转日期序列号。

    增加『合理年份区间』守卫（1970–2100）：只有转换结果落在该区间才认为是
    真日期（如成立日期 39120 → 2007-02-07）；否则原样保留数字，避免把
    邮政编码 453400 → 3141-05-13、缴纳社保 757 → 1902-01-26、
    公司总人数 2600 → 1907-02-12 等纯数字误转成荒谬日期。
    新代码请优先用 cell_str(v, fmt)。
    """
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        d = excel_serial_to_date(v)
        if d:
            try:
                year = int(d[:4])
                if 1970 <= year <= 2100:
                    return d
            except Exception:
                pass
            # 超出合理年份区间 → 视为普通数字，落到下方原样返回
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip()


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path
