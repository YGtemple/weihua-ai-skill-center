#!/usr/bin/env python3
"""检查集团总裁发言稿纯文本的形式质量；不核验事实真伪。"""

import argparse
import re
import sys
from pathlib import Path


CN_NUMERALS = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}

SCENES = ["general", "annual-commendation", "quarterly", "half-year", "year-end", "mobilization"]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path, help="UTF-8 纯文本发言稿")
    parser.add_argument("--scene", choices=SCENES, default="general", help="会议场景，默认 general")
    parser.add_argument("--min-chars", type=int, default=0, help="最少非空白字符数，默认不限制")
    parser.add_argument("--max-chars", type=int, default=0, help="最多非空白字符数，0 表示不限制")
    parser.add_argument("--allow-placeholders", action="store_true", help="允许待核占位符，但仍列出")
    parser.add_argument("--free-structure", action="store_true", help="不检查年度表彰会五部分默认结构")
    args = parser.parse_args()

    try:
        text = args.file.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        print(f"错误：无法读取 {args.file}: {exc}", file=sys.stderr)
        return 2

    errors = []
    warnings = []
    char_count = len(re.sub(r"\s+", "", text))
    print(f"场景：{args.scene}")
    print(f"非空白字符数：{char_count}")

    if args.min_chars and char_count < args.min_chars:
        warnings.append(f"篇幅少于设定的 {args.min_chars} 字符")
    if args.max_chars and char_count > args.max_chars:
        warnings.append(f"篇幅超过设定的 {args.max_chars} 字符")

    markdown_lines = [line for line in text.splitlines() if re.match(r"^\s{0,3}#{1,6}\s+", line)]
    if markdown_lines:
        errors.append(f"含 {len(markdown_lines)} 行 Markdown 标题，请改成自然讲话层级")

    editorial_patterns = [
        r"以下是(?:讲话|讲话稿|发言稿)",
        r"作为(?:AI|人工智能)",
        r"TODO",
        r"待完善",
        r"本文将从",
        r"综上所述",
    ]
    for pattern in editorial_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            errors.append(f"含编辑或生成痕迹：{pattern}")

    drafting_notes = re.findall(
        r"[（(][^）)\n]*(?:前排|临场|口语|这里补充|可尝试|举例：|待补|待核)[^）)\n]*[）)]",
        text,
    )
    if drafting_notes:
        warnings.append(f"发现 {len(drafting_notes)} 处疑似起草括注，请改为正式正文或删除")

    placeholder_pattern = r"(?:[\[【（(]\s*待(?:核|补|确认|提供)[^\]】）)\n]*[\]】）)]|XX+|××+)"
    placeholders = re.findall(placeholder_pattern, text)
    if placeholders:
        message = f"发现 {len(placeholders)} 处待核或占位标记"
        if args.allow_placeholders:
            warnings.append(message + "（本次允许）")
        else:
            errors.append(message + "；正式净稿须清零")

    if not re.search(r"^(?:同志们|同事们|各位同事|全体同事|各位领导|各位嘉宾)", text, re.MULTILINE):
        warnings.append("未检出常见现场称谓，请按实际听众核对")
    if "集团总裁" not in text:
        warnings.append("未检出“集团总裁”署名，请核对标题与署名")

    if args.scene == "annual-commendation":
        if "表彰" not in text:
            warnings.append("年度表彰会未检出“表彰”，请确认会议定位")
        if "祝贺" not in text:
            warnings.append("年度表彰会未检出“祝贺”，请确认是否向受表彰者致意")
        if not re.search(r"先进|优秀|模范|奋斗者", text):
            warnings.append("未检出先进、优秀、模范或奋斗者等表彰对象，请核对表彰主线")

    top_sections = []
    for line in text.splitlines():
        match = re.match(r"^\s*([一二三四五六七八九十])[、，]\s*\S+", line)
        if match:
            top_sections.append(CN_NUMERALS[match.group(1)])
    if top_sections:
        expected = list(range(1, len(top_sections) + 1))
        if top_sections != expected:
            warnings.append(f"中文数字层级可能跳号或重复：{top_sections}，请核对编号")

    if args.scene == "annual-commendation" and not args.free_structure:
        missing = [str(i) for i in range(1, 6) if i not in top_sections]
        if missing:
            warnings.append(
                "未检出完整的五部分默认层次，缺少编号：" + "、".join(missing)
                + "；如为有意合并可忽略或加 --free-structure"
            )

    if re.search(r"个百分(?!点)", text):
        errors.append("发现“个百分”表述，比例差值应核对并写为“个百分点”")
    if re.search(r"数\d+(?:万元|亿元|台|项|人)", text):
        warnings.append("发现中文约数与阿拉伯数字混用，请统一为准确数字或规范中文约数")
    if re.search(r"!{2,}|！{2,}", text):
        warnings.append("发现连续感叹号，正式讲话建议克制使用")

    strong_claims = re.findall(r"创历史新高|历史最高|全面完成|圆满完成|行业领先|全球首台套|填补空白|零事故|100%", text)
    if strong_claims:
        print("强判断提醒：" + "、".join(sorted(set(strong_claims))) + "（须逐项核验直接证据）")

    paragraphs = [re.sub(r"\s+", "", p) for p in re.split(r"\n\s*\n", text)]
    seen = set()
    duplicates = set()
    for paragraph in paragraphs:
        if len(paragraph) >= 30:
            if paragraph in seen:
                duplicates.add(paragraph[:24])
            seen.add(paragraph)
    if duplicates:
        warnings.append(f"发现 {len(duplicates)} 个重复长段落：" + "；".join(sorted(duplicates)))

    number_count = len(re.findall(r"\d+(?:\.\d+)?\s*(?:%|％|亿元|万元|万|台|项|人|天|年|月)?", text))
    if number_count:
        print(f"数字表达：{number_count} 处（请另行核对期间、单位、口径、来源和派生关系）")

    for item in errors:
        print("错误：" + item)
    for item in warnings:
        print("提醒：" + item)
    if not errors and not warnings:
        print("形式检查通过；仍须人工完成事实与语感审校。")

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
