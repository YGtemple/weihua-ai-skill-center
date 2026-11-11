#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""宣传文案机检脚本：字数/口播时长/风险词/长句/标点。

用法：
  python check_copy.py --file 稿件.txt
  python check_copy.py --text "稿件正文"
  python check_copy.py --file 稿件.txt --duration 27.4

只做提示，不做判定：命中不等于违规，需人工看语境。

架构说明（标准与代码分离）：
- 风险词库、豁免上下文在 config/risk-terms.json，加词/改类别只改那里。
- 语速档、默认语速、长句阈值、超载/不足系数在 config/video-profile.json，
  CLI --limit 仍可覆盖配置里的长句阈值。
- 配置缺失或损坏时本脚本会给出清晰错误并退出，不会静默用默认值。
"""

import argparse
import json
import os
import re
import sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(SKILL_DIR, "config")


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


RISK_CONFIG = load_config("risk-terms.json")
VIDEO_PROFILE = load_config("video-profile.json")

# 风险词库（顺序即扫描顺序），外置到 config/risk-terms.json
RISK_TERMS = [(t["term"], t["category"], t["note"]) for t in RISK_CONFIG["terms"]]
EXCLUDE_PATTERNS = RISK_CONFIG.get("exclude_patterns", {})
SPECIAL_TERMS = [(t["term"], t["category"], t["note"]) for t in RISK_CONFIG.get("special_terms", [])]

# 机械参数，外置到 config/video-profile.json
SPEECH_RATES = VIDEO_PROFILE["speech_rates"]
DEFAULT_RATE = VIDEO_PROFILE.get("default_rate", SPEECH_RATES[1])
LONG_SENTENCE_LIMIT = int(VIDEO_PROFILE.get("long_sentence_limit", 40))
_DC = VIDEO_PROFILE.get("duration_check", {})
OVERLOAD_RATIO = float(_DC.get("overload_ratio", 1.2))
UNDERLOAD_RATIO = float(_DC.get("underload_ratio", 0.7))


def char_count(text):
    """返回 (汉字数, 其他有效字符数, 总字数估计)。"""
    han = re.findall(r"[\u4e00-\u9fff]", text)
    others = re.findall(r"[A-Za-z0-9]", text)
    return len(han), len(others), len(han) + round(len(others) * 0.6)


def find_context(text, start, end, width=14):
    left = max(0, start - width)
    right = min(len(text), end + width)
    snippet = text[left:right].replace("\n", " ")
    prefix = "…" if left > 0 else ""
    suffix = "…" if right < len(text) else ""
    return f"{prefix}{snippet}{suffix}"


def scan_terms(text):
    hits = []
    consumed = set()

    for term, cat, note in RISK_TERMS:
        for m in re.finditer(re.escape(term), text):
            pos = m.start()
            if pos in consumed:
                continue
            excluded = False
            for pat in EXCLUDE_PATTERNS.get(term, []):
                if text.startswith(pat, pos) or text[max(0, pos - 4):pos + len(pat)].endswith(pat):
                    excluded = True
                    break
            if excluded:
                continue
            consumed.add(pos)
            hits.append({
                "term": term, "cat": cat, "note": note,
                "pos": pos + 1,
                "ctx": find_context(text, m.start(), m.end()),
            })

    for term, cat, note in SPECIAL_TERMS:
        for m in re.finditer(re.escape(term), text):
            pos = m.start()
            if pos in consumed:
                continue
            excluded = False
            for pat in EXCLUDE_PATTERNS.get(term, []):
                if text.startswith(pat, pos):
                    excluded = True
                    break
            if excluded:
                continue
            consumed.add(pos)
            hits.append({
                "term": term, "cat": cat, "note": note,
                "pos": pos + 1,
                "ctx": find_context(text, m.start(), m.end()),
            })

    hits.sort(key=lambda h: h["pos"])
    return hits


def long_sentences(text, limit=40):
    parts = re.split(r"[。！？；\n]+", text)
    out = []
    for p in parts:
        s = p.strip()
        if len(s) > limit:
            out.append((len(s), s))
    return out


def punctuation_issues(text):
    issues = []
    if re.search(r"[\u4e00-\u9fff]，|[\u4e00-\u9fff]\,", text):
        if re.search(r"[\u4e00-\u9fff],", text):
            issues.append("中文语境中出现半角逗号「,」")
    if re.search(r"[\u4e00-\u9fff]\(", text):
        issues.append("中文语境中出现半角括号「()」")
    if re.search(r"[\u4e00-\u9fff]\.", text):
        issues.append("中文语境中出现半角句点「.」")
    if "  " in text:
        issues.append("存在连续空格（复制粘贴导致）")
    return issues


def number_issues(text):
    issues = []
    # 缺年份的时间表述：出现「X月X日」但全文没有 20xx 年
    if re.search(r"\d{1,2}\s*月\s*\d{1,2}\s*日", text) and not re.search(r"20\d{2}\s*年", text):
        issues.append("出现「X月X日」但全文未写年份（对外稿建议补全年份）")
    if re.search(r"\d+\.\d+\s*日", text):
        issues.append("日期写成小数样式（如 11.6 日），口播与字幕均不友好")
    if re.search(r"[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日", text):
        issues.append("日期含年份，请确认是否为发布稿需要（活动稿常只需月日，通稿建议完整）")
    return issues


def load_text(args):
    if args.text:
        return args.text
    with open(args.file, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    sys.exit("无法识别文件编码，请另存为 UTF-8 后重试")


def main():
    ap = argparse.ArgumentParser(description="宣传文案机检")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--file", help="稿件文本文件路径")
    g.add_argument("--text", help="直接传入稿件正文")
    ap.add_argument("--duration", type=float, default=None, help="成片时长（秒），用于核算字数是否装得下")
    ap.add_argument("--limit", type=int, default=None,
                    help=f"长句阈值（默认取 config/video-profile.json 的 {LONG_SENTENCE_LIMIT} 字，可覆盖）")
    args = ap.parse_args()

    limit = args.limit if args.limit is not None else LONG_SENTENCE_LIMIT

    text = load_text(args).strip()
    if not text:
        sys.exit("稿件为空")

    han, others, total = char_count(text)

    print("=" * 56)
    print("文案机检报告")
    print("=" * 56)
    print(f"总字符 {len(text)}｜有效字数 {total}（汉字 {han}，字母数字折 {total - han}）")
    rate_segs = "｜".join(f"{r:g}字/秒 → {total / r:.1f}s" for r in SPEECH_RATES)
    print(f"预估口播时长：{rate_segs}")

    if args.duration:
        d = args.duration
        low, high = round(d * min(SPEECH_RATES)), round(d * max(SPEECH_RATES))
        mid = round(d * DEFAULT_RATE)
        print(f"\n成片时长 {d:.1f}s 承载量：{low}–{high} 字（按 {DEFAULT_RATE:g} 字/秒取 {mid} 字）")
        if total > high * OVERLOAD_RATIO:
            print(f"→ 判定：超载 {(total / high - 1) * 100:.0f}%。口播装不下，需删至 {mid} 字内或延长时长。")
        elif total > high:
            print("→ 判定：略超上限，需压缩或加快语速（注意字幕挤压）。")
        elif total < low * UNDERLOAD_RATIO:
            print("→ 判定：严重偏少，画面会空转，建议补信息点。")
        else:
            print("→ 判定：字数与时长匹配正常。")

    hits = scan_terms(text)
    print(f"\n风险词命中 {len(hits)} 处")
    if hits:
        for h in hits:
            print(f"  [{h['cat']}] 「{h['term']}」第{h['pos']}字｜{h['note']}")
            print(f"      语境：{h['ctx']}")
    else:
        print("  未命中内置词库（不代表没有风险，仍需人工看关系表述与事实）")

    ls = long_sentences(text, limit)
    print(f"\n超长句（>{limit} 字）{len(ls)} 处")
    for n, s in ls:
        print(f"  {n} 字：{s[:70]}{'…' if len(s) > 70 else ''}")

    pi = punctuation_issues(text)
    ni = number_issues(text)
    if pi or ni:
        print("\n格式与数字提示")
        for i in pi + ni:
            print(f"  - {i}")

    print("\n提示：以上为机械扫描，命中不等于违规。引用政策原文、客观事实陈述可豁免；")
    print("      关系表述（主办/受邀）、人物职务、名单一致性需人工核对。")


if __name__ == "__main__":
    main()
