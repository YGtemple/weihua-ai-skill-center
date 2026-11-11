#!/usr/bin/env python3
"""检查董事长发言稿纯文本的形式质量；不核验事实真伪。"""

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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path, help="UTF-8 纯文本发言稿")
    parser.add_argument(
        "--scene",
        choices=["general", "annual-commendation", "quarterly", "half-year", "year-end", "mobilization"],
        default="general",
        help="会议场景，默认 general；仅影响场景提醒，不限制写法",
    )
    parser.add_argument("--min-chars", type=int, default=0, help="最少非空白字符数，默认不限制")
    parser.add_argument("--max-chars", type=int, default=0, help="最多非空白字符数，0 表示不限制")
    parser.add_argument("--allow-placeholders", action="store_true", help="允许待核占位符，但仍列出")
    parser.add_argument(
        "--free-structure",
        action="store_true",
        help="不检查年度表彰会的三部分或五部分默认结构",
    )
    args = parser.parse_args()

    try:
        text = args.file.read_text(encoding="utf-8-sig")
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
        errors.append(f"含 {len(markdown_lines)} 行 Markdown 标题，请改成讲话稿自然层级")

    editorial_patterns = [
        r"这里重点提到",
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

    top_sections = []
    for line in text.splitlines():
        # 识别"一、""二、"格式
        match = re.match(r"^\s*([一二三四五六七八九十])[、，]\s*\S+", line)
        if match:
            top_sections.append(CN_NUMERALS[match.group(1)])
            continue
        # 识别"第一部分""第二部分"格式
        match = re.match(r"^\s*第([一二三四五六七八九十])(?:部分|个方面)[、，：:]?\s*\S*", line)
        if match:
            top_sections.append(CN_NUMERALS[match.group(1)])
    if top_sections and top_sections != list(range(1, len(top_sections) + 1)):
        warnings.append(f"中文数字层级可能跳号或重复：{top_sections}，请核对编号")

    if args.scene == "annual-commendation":
        if "表彰" not in text:
            warnings.append("年度表彰会未检出“表彰”，请确认会议定位")
        if "祝贺" not in text:
            warnings.append("年度表彰会未检出“祝贺”，请确认是否向受表彰者致意")
        if not re.search(r"先进|优秀|模范|奋斗者", text):
            warnings.append("未检出先进、优秀、模范或奋斗者等表彰对象，请核对表彰主线")
        if not args.free_structure:
            has_three = (
                all(i in top_sections for i in (1, 2, 3))
                or all(
                    re.search(pattern, text, re.MULTILINE)
                    for pattern in (r"第一个方面", r"第二个方面", r"第三个方面")
                )
            )
            has_five = has_three and all(i in top_sections for i in (4, 5))
            if not (has_three or has_five):
                warnings.append(
                    "未检出三部分或五部分默认层次；如为有意设计可忽略或加 --free-structure"
                )

    if args.scene == "year-end":
        if not re.search(r"大势|形势|大局", text):
            warnings.append("收官会未检出“大势/形势/大局”类判断，请确认是否有外部环境分析")
        if not re.search(r"收官|冲刺|最后", text):
            warnings.append("收官会未检出“收官/冲刺/最后”类表述，请确认会议定位")
        if not re.search(r"任务|要求|部署", text):
            warnings.append("收官会未检出“任务/要求/部署”类行动指向，请确认是否有下阶段安排")

    if re.search(r"个百分(?!点)", text):
        errors.append("发现“个百分”表述，比例差值应核对并写为“个百分点”")
    if re.search(r"数\d+(?:万元|亿元|台|项|人)", text):
        warnings.append("发现中文约数与阿拉伯数字混用，请统一为准确数字或规范中文约数")
    if re.search(r"!{2,}|！{2,}", text):
        warnings.append("发现连续感叹号，正式讲话建议克制使用")

    strong_claims = re.findall(
        r"创历史新高|全面完成|圆满完成|行业领先|全球首台套|填补空白|零事故|100%",
        text,
    )
    if strong_claims:
        print("强判断提醒：" + "、".join(sorted(set(strong_claims))) + "（须逐项核验直接证据）")

    # 身份自检：只拦截说话人自我身份污染，不误伤议程中正常提及总裁
    identity_patterns = [
        r"我作为总裁",
        r"我以总裁名义",
        r"我是总裁",
        r"署名.*总裁",
        r"我代表总裁",
    ]
    identity_hits = []
    for pattern in identity_patterns:
        if re.search(pattern, text):
            identity_hits.append(pattern)
    if identity_hits:
        errors.append(f"身份污染：说话人自称总裁，本 Skill 仅用于董事长发言稿（{', '.join(identity_hits)}）")

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
