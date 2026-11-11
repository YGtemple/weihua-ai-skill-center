# -*- coding: utf-8 -*-
"""一次性把 11 个技能包里的资料库根目录换成你们实际的路径。

用法（在「卫华技能中心交付」目录下执行）：

    python "_配置/set_library_path.py" --root "D:\\投标资料库" --dry-run
    python "_配置/set_library_path.py" --root "D:\\投标资料库"

只改文本里的路径字符串，不动任何台账、不改脚本逻辑。
"""
import argparse
import os
import sys

DEFAULT_OLD = r"E:\卫华供应商入围资料库"
TEXT_EXT = {".md", ".py", ".json", ".txt", ".csv", ".yaml", ".yml"}


def variants(path):
    """同一个路径在文件里的三种写法：原生、双反斜杠、正斜杠。"""
    return [
        path,
        path.replace("\\", "\\\\"),
        path.replace("\\", "/"),
    ]


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(here)
    packages = os.path.join(root_dir, "卫华AI技能包")

    parser = argparse.ArgumentParser(description="替换技能包内的资料库根目录")
    parser.add_argument("--root", required=True, help="新的资料库根目录，例如 D:\\投标资料库")
    parser.add_argument("--old", default=DEFAULT_OLD, help="原来的根目录（默认 E:\\卫华供应商入围资料库）")
    parser.add_argument("--dry-run", action="store_true", help="只统计，不写盘")
    args = parser.parse_args()

    if not os.path.isdir(packages):
        sys.exit("找不到技能包目录：%s" % packages)

    old_forms = variants(args.old)
    new_forms = variants(args.root)

    changed = []
    total_hits = 0

    for base, _dirs, files in os.walk(packages):
        for name in files:
            if os.path.splitext(name)[1].lower() not in TEXT_EXT:
                continue
            path = os.path.join(base, name)
            with open(path, encoding="utf-8") as f:
                text = f.read()
            hits = 0
            for old, new in zip(old_forms, new_forms):
                hits += text.count(old)
                text = text.replace(old, new)
            if hits:
                total_hits += hits
                changed.append((os.path.relpath(path, root_dir), hits))
                if not args.dry_run:
                    with open(path, "w", encoding="utf-8", newline="\n") as f:
                        f.write(text)

    action = "将修改" if args.dry_run else "已修改"
    print("%s %d 个文件，共 %d 处：" % (action, len(changed), total_hits))
    for rel, hits in changed:
        print("  %s  (%d 处)" % (rel, hits))
    if args.dry_run:
        print("\n这是预演，没有写盘。去掉 --dry-run 再执行一次即生效。")


if __name__ == "__main__":
    main()
