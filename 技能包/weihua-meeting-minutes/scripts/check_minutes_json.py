#!/usr/bin/env python3
"""Validate Dafang meeting-minutes JSON, including optional source coverage."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

REQUIRED_FIELDS = ("会议主题", "会议时间", "会议地点", "主持人", "参会人员", "整理人", "审核人")
BLOCK_TYPES = {"h2", "h3", "department", "reply", "body", "paragraph", "item", "page_break"}
# 事实块：必须能追溯到已声明来源（block.source_ids 或 section.default_source_ids）
FACT_BLOCK_TYPES = {"body", "paragraph", "item"}
# 结构标题块：不强制来源（h2/h3/department/reply 本身是骨架，不是事实陈述）
STRUCTURE_BLOCK_TYPES = {"h2", "h3", "department", "reply"}

# 占位符：覆盖常见写法，但不误伤中文省略号“……”（U+2026）和正常句号。
PLACEHOLDER_RE = re.compile(
    r"×{2,}"
    r"|【\s*待(?:核|定|补)"
    r"|\[\s*待(?:核|定|补)"
    r"|（\s*待(?:核|定|补)"
    r"|\(\s*待(?:核|定|补)"
    r"|<\s*待(?:核|定|补)"
    r"|TODO"
    r"|TBD"
    r"|N/?A\b"
    r"|xxx",
    re.IGNORECASE,
)

# 集团总裁 / 集团董事长章节：用于检测“挂名但无内容”的空领导章节。
LEADER_HEADING_RE = re.compile(r"集团\s*(?:总裁|董事长)")


def _has_placeholder(text: str) -> bool:
    return bool(text and PLACEHOLDER_RE.search(text))


def validate(data: dict, allow_placeholders: bool = False) -> tuple[list[str], dict]:
    errors: list[str] = []
    if not isinstance(data, dict):
        return ["根节点必须是JSON对象"], {}

    meeting_type = data.get("meeting_type")
    if meeting_type not in {"monthly", "quarterly", "special"}:
        errors.append("meeting_type 必须为 monthly、quarterly 或 special")

    for key in ("company", "issuer"):
        if not isinstance(data.get(key), str) or not data.get(key, "").strip():
            errors.append(f"{key} 必须为非空字符串")

    # ---- 首页字段 ----
    fields = data.get("fields")
    labels: list[str] = []
    if not isinstance(fields, list) or not fields:
        errors.append("fields 必须为非空数组")
        fields = []
    for i, field in enumerate(fields):
        if not isinstance(field, dict):
            errors.append(f"fields[{i}] 必须为对象")
            continue
        label, value = field.get("label"), field.get("value")
        if not isinstance(label, str) or not label.strip():
            errors.append(f"fields[{i}].label 无效")
            continue
        labels.append(label)
        if not isinstance(value, str) or not value.strip():
            errors.append(f"字段“{label}”的值不能为空")
        elif not allow_placeholders and _has_placeholder(value):
            errors.append(f"字段“{label}”仍含未解释占位符")
    for label in REQUIRED_FIELDS:
        if label not in labels:
            errors.append(f"缺少必需字段：{label}")

    # ---- 议程 ----
    agenda = data.get("agenda", [])
    if not isinstance(agenda, list) or any(not isinstance(x, str) or not x.strip() for x in agenda):
        errors.append("agenda 必须为字符串数组")
        agenda = []
    if meeting_type == "quarterly" and not agenda:
        errors.append("季度大会必须填写会议整体议程")

    # ---- 来源登记 ----
    sources = data.get("sources", [])
    source_required: dict[str, bool] = {}
    if not isinstance(sources, list):
        errors.append("sources 必须为数组")
        sources = []
    for i, source in enumerate(sources):
        if not isinstance(source, dict) or not isinstance(source.get("id"), str) or not source["id"].strip():
            errors.append(f"sources[{i}] 必须含非空 id")
            continue
        sid = source["id"]
        if sid in source_required:
            errors.append(f"来源ID重复：{sid}")
        source_required[sid] = bool(source.get("required", True))

    used: set[str] = set()
    sections = data.get("sections")
    block_count = 0
    if not isinstance(sections, list) or not sections:
        errors.append("sections 必须为非空数组")
        sections = []
    headings: list[str] = []

    for si, section in enumerate(sections):
        if not isinstance(section, dict):
            errors.append(f"sections[{si}] 必须为对象")
            continue
        heading = section.get("heading")
        if not isinstance(heading, str) or not heading.strip():
            errors.append(f"sections[{si}].heading 无效")
            heading = ""
        else:
            headings.append(heading)
            if not allow_placeholders and _has_placeholder(heading):
                errors.append(f"章节标题仍含占位符：{heading}")

        # ---- section 级 default_source_ids（块未显式挂来源时继承） ----
        section_default = section.get("default_source_ids", [])
        if section_default is None:
            section_default = []
        if not isinstance(section_default, list) or any(not isinstance(x, str) for x in section_default):
            errors.append(f"章节“{heading}”的 default_source_ids 必须为字符串数组")
            section_default = []
        for sid in section_default:
            if sources and sid not in source_required:
                errors.append(f"章节“{heading}”声明的 default_source_ids 引用未知来源：{sid}")
            used.add(sid)

        blocks = section.get("blocks")
        if not isinstance(blocks, list):
            errors.append(f"章节“{heading}”的 blocks 必须为数组")
            continue

        # ---- 空领导章节检测：heading 含“集团总裁/集团董事长”但无内容 ----
        # 正式净稿（allow_placeholders=False）：blocks 为空或全是占位符都报错。
        # 模板测试（allow_placeholders=True）：只拦 blocks 完全为空（结构缺失），
        # 纯占位符块视为“模板待填”，不在这里报错（占位符由外层规则统一报）。
        is_leader_section = bool(heading and LEADER_HEADING_RE.search(heading))
        non_empty_blocks = [
            b for b in blocks
            if isinstance(b, dict)
            and b.get("type") != "page_break"
            and isinstance(b.get("text"), str)
            and b["text"].strip()
        ]
        if allow_placeholders:
            has_content = bool(non_empty_blocks)
        else:
            has_content = any(not _has_placeholder(b.get("text", "")) for b in non_empty_blocks)
        if is_leader_section and not has_content:
            errors.append(
                f"章节“{heading}”为空或仅含占位符；若集团总裁/董事长本次未出席或未发言，"
                f"应整段删除该章节，不要保留空标题"
            )

        previous_text = None
        # 空部门检测：遇到 department 后向后扫描，直到下一个 department 或章节结束，
        # 中间必须出现至少一个事实块（item/body/paragraph）。
        pending_department: str | None = None

        for bi, block in enumerate(blocks):
            block_count += 1
            if not isinstance(block, dict):
                errors.append(f"{heading}.blocks[{bi}] 必须为对象")
                continue
            btype = block.get("type")
            if btype not in BLOCK_TYPES:
                errors.append(f"{heading}.blocks[{bi}] 未知类型：{btype}")
            text = block.get("text", "")

            if btype == "page_break":
                continue
            if not isinstance(text, str) or not text.strip():
                errors.append(f"{heading}.blocks[{bi}] 文本不能为空")
                continue
            if not allow_placeholders and _has_placeholder(text):
                errors.append(f"正文仍含占位符：{text[:30]}")
            if text == previous_text and btype not in {"department", "reply"}:
                errors.append(f"相邻正文重复：{text[:30]}")
            previous_text = text

            if btype == "item":
                level = block.get("level", 0)
                if not isinstance(level, int) or isinstance(level, bool) or not 0 <= level <= 3:
                    errors.append(f"条目层级必须为0—3整数：{text[:30]}")

            # ---- 解析最终生效的 source_ids：显式优先，缺省继承 section.default_source_ids ----
            block_sids = block.get("source_ids")
            if block_sids is None:
                effective_sids: list[str] = list(section_default)
            else:
                if not isinstance(block_sids, list) or any(not isinstance(x, str) for x in block_sids):
                    errors.append(f"{heading}.blocks[{bi}].source_ids 必须为字符串数组")
                    effective_sids = []
                else:
                    effective_sids = list(block_sids)

            for sid in effective_sids:
                used.add(sid)
                if sources and sid not in source_required:
                    errors.append(f"正文引用未知来源ID：{sid}")

            # ---- 事实块必须有来源 ----
            if btype in FACT_BLOCK_TYPES and not effective_sids:
                errors.append(
                    f"事实块缺少来源标注：“{text[:30]}”；"
                    f"请在该 block 上加 source_ids，或在所属 section 上加 default_source_ids"
                )

            # ---- 空部门检测状态机 ----
            if btype == "department":
                if pending_department is not None:
                    errors.append(
                        f"空部门标题：“{pending_department}”下没有任何 item/body/回复内容；"
                        f"若该部门本次未发言，应整段删除"
                    )
                pending_department = text
            elif btype in FACT_BLOCK_TYPES:
                # 事实块出现，说明上一个 department 有内容
                pending_department = None
            # reply 是结构标题，本身不算事实内容，但也不清空 pending_department，
            # 因为 reply 后面通常还应跟一个 body 才构成完整回复。

        # 章节结束，最后一个 department 仍可能悬空
        if pending_department is not None:
            errors.append(
                f"空部门标题：“{pending_department}”下没有任何 item/body/回复内容；"
                f"若该部门本次未发言，应整段删除"
            )

    # ---- 月度特有：前两章关键词 ----
    if meeting_type == "monthly" and len(headings) >= 2:
        expected = ("主持人通报", "各部门主管")
        for i, keyword in enumerate(expected):
            if keyword not in headings[i]:
                errors.append(f"月度第{i + 1}部分应包含“{keyword}”，当前为：{headings[i]}")

    # ---- 季度特有：议程数量与章节数量一致（只做数量比对，不做语义猜测） ----
    if meeting_type == "quarterly" and agenda and headings:
        if len(agenda) != len(headings):
            errors.append(
                f"季度大会议程数量（{len(agenda)}条）与正文一级章节数量（{len(headings)}个）不一致；"
                f"请按实际议程增删 agenda 或 sections"
            )

    # ---- 排除来源 ----
    excluded: set[str] = set()
    exclusions = data.get("excluded_sources", [])
    if not isinstance(exclusions, list):
        errors.append("excluded_sources 必须为数组")
        exclusions = []
    for i, item in enumerate(exclusions):
        if not isinstance(item, dict) or not isinstance(item.get("id"), str):
            errors.append(f"excluded_sources[{i}] 必须含来源 id")
            continue
        sid = item["id"]
        if sources and sid not in source_required:
            errors.append(f"排除清单引用未知来源ID：{sid}")
        if sid in excluded:
            errors.append(f"排除来源ID重复：{sid}")
        excluded.add(sid)
        if not isinstance(item.get("reason"), str) or not item["reason"].strip():
            errors.append(f"排除来源 {sid} 必须说明原因")
    for sid in sorted(used & excluded):
        errors.append(f"来源 {sid} 同时被采用和排除")

    uncovered = []
    if sources:
        for sid, required in source_required.items():
            if required and sid not in used and sid not in excluded:
                uncovered.append(sid)
                errors.append(f"必需来源 {sid} 未被正文采用，也未说明排除原因")

    stats = {
        "sections": len(sections),
        "blocks": block_count,
        "sources": len(source_required),
        "used_sources": len(used),
        "excluded_sources": len(excluded),
        "uncovered_sources": uncovered,
    }
    return errors, stats


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("json_file", type=Path)
    parser.add_argument("--allow-placeholders", action="store_true", help="仅用于测试含××占位的模板")
    args = parser.parse_args()
    try:
        # utf-8-sig 自动兼容带 BOM（Windows 记事本/PowerShell 默认）与无 BOM 的 JSON。
        data = json.loads(args.json_file.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"读取JSON失败：{exc}", file=sys.stderr)
        return 2
    errors, stats = validate(data, args.allow_placeholders)
    if errors:
        print("校验失败：", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    print("校验通过：" + "，".join(f"{k}={v}" for k, v in stats.items() if k != "uncovered_sources"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
