#!/usr/bin/env python3
"""chairman-speech Skill 结构快速校验。

无第三方依赖，仅使用标准库。检查项：
1. SKILL.md frontmatter 存在且 name 与目录名一致；
2. SKILL.md 中引用的本地 Markdown 文件全部存在；
3. scripts/speech_lint.py 可被 Python 编译解析；
4. references/ 下被引用的文件无悬空链接。

用法：
    python scripts/quick_validate.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path


def skill_root() -> Path:
    """返回本 Skill 根目录（scripts/ 的上一级）。"""
    return Path(__file__).resolve().parent.parent


def fail(msg: str) -> None:
    print(f"错误：{msg}", file=sys.stderr)


def pass_(msg: str) -> None:
    print(f"通过：{msg}")


def check_frontmatter(root: Path) -> bool:
    """校验 SKILL.md 的 frontmatter 存在且 name 字段与目录名一致。"""
    skill_md = root / "SKILL.md"
    if not skill_md.is_file():
        fail("SKILL.md 不存在")
        return False

    text = skill_md.read_text(encoding="utf-8-sig")
    m = re.match(r"^---\n(.*?)\n---", text, re.DOTALL)
    if not m:
        fail("SKILL.md 缺少 frontmatter（--- 包裹的 YAML 头）")
        return False

    fm = m.group(1)
    name_m = re.search(r"^name:\s*(\S+)\s*$", fm, re.MULTILINE)
    if not name_m:
        fail("frontmatter 中缺少 name 字段")
        return False

    name_value = name_m.group(1).strip().strip('"').strip("'")
    dir_name = root.name
    if name_value != dir_name:
        fail(f"frontmatter name='{name_value}' 与目录名 '{dir_name}' 不一致")
        return False

    pass_(f"frontmatter name='{name_value}' 与目录名一致")
    return True


def check_local_refs(root: Path) -> bool:
    """检查 SKILL.md 和 references/ 下所有 .md 中引用的本地 .md 文件是否存在。"""
    md_files = [root / "SKILL.md"]
    md_files.extend(sorted((root / "references").glob("*.md")))

    # 匹配 [xxx](path.md) 或 [xxx](references/path.md) 形式的本地链接
    link_re = re.compile(r"\[[^\]]*\]\(([^)]+\.md)\)")
    # 匹配反引号中的相对路径引用，如 references/drafting-patterns.md
    backtick_re = re.compile(r"`((?:references|scripts|assets)/[^\s`]+)`")

    # 已知的跨 Skill 引用：路径形式与本目录相同，但实际指向其他 Skill 目录，
    # 例如内置 word Skill 的 Windows 兼容文件。这些路径不在本目录下检查。
    cross_skill_refs = {
        "references/workflows/windows-compat.md",
    }

    missing: list[str] = []
    checked: set[Path] = set()

    for md in md_files:
        if not md.is_file():
            continue
        text = md.read_text(encoding="utf-8-sig")
        base_dir = md.parent

        for match in link_re.finditer(text):
            ref = match.group(1).strip()
            # 跳过 URL
            if "://" in ref or ref.startswith(("http", "mailto:")):
                continue
            target = (base_dir / ref).resolve()
            checked.add(target)
            if not target.is_file():
                missing.append(f"{md.name} -> {ref}")

        for match in backtick_re.finditer(text):
            ref = match.group(1).strip()
            if ref in cross_skill_refs:
                continue  # 跨 Skill 引用，不在本目录下检查
            target = (root / ref).resolve()
            checked.add(target)
            if not target.is_file():
                missing.append(f"{md.name} -> `{ref}`")

    if missing:
        for item in missing:
            fail(f"本地引用不存在：{item}")
        return False

    pass_(f"共校验 {len(checked)} 个本地引用目标，全部存在")
    return True


def check_lint_compiles(root: Path) -> bool:
    """检查 speech_lint.py 可被 Python 编译解析。"""
    lint = root / "scripts" / "speech_lint.py"
    if not lint.is_file():
        fail("scripts/speech_lint.py 不存在")
        return False

    try:
        source = lint.read_text(encoding="utf-8")
        compile(source, str(lint), "exec")
    except SyntaxError as exc:
        fail(f"speech_lint.py 语法错误：第 {exc.lineno} 行 {exc.msg}")
        return False
    except OSError as exc:
        fail(f"无法读取 speech_lint.py：{exc}")
        return False

    pass_("speech_lint.py 可正常编译解析")
    return True


def check_references_dir(root: Path) -> bool:
    """references/ 目录应存在且至少包含 5 个 .md 文件。"""
    refs = root / "references"
    if not refs.is_dir():
        fail("references/ 目录不存在")
        return False
    md_count = len(list(refs.glob("*.md")))
    if md_count < 5:
        fail(f"references/ 下仅 {md_count} 个 .md 文件，疑似缺失模块")
        return False
    pass_(f"references/ 下有 {md_count} 个 .md 文件")
    return True


def main() -> int:
    root = skill_root()
    print(f"校验目录：{root}")
    print("-" * 50)

    results = [
        check_frontmatter(root),
        check_references_dir(root),
        check_local_refs(root),
        check_lint_compiles(root),
    ]

    print("-" * 50)
    if all(results):
        print("全部结构校验通过。")
        return 0
    print(f"共 {results.count(False)} 项校验失败。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
