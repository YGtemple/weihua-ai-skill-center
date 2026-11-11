# -*- coding: utf-8 -*-
"""
按 01_报名 / 02_投标 子文件夹归集合规资料副本（v1.4）。

硬性规则：
- 仅 ✅ 合规项对应文件按阶段归类；
- 仅复制（shutil.copy2），绝不移动、修改、删除原文件；
- 现场备查、中标后提交的资料不纳入本次归集；
- 源文件不存在时报错并跳过该项，禁止伪造路径。

用法：
  python collect_compliant_files.py --items items.json --root 输出根目录
  python collect_compliant_files.py --items items.json --root 输出根目录 --dry-run
    --items 审核数据 JSON（含 collected_files 段）
    --root  本次招标输出根目录（含"项目名_项目编号"或日期的目录）
"""
import argparse
import datetime as dt
import json
import shutil
import sys
from pathlib import Path

DISCLAIMER = ("本归集仅复制合规资料副本，原文件保留原位未动。"
              "AI仅辅助分析，禁止直接对外提交；所有归集结果必须人工抽检。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--items", required=True,
                    help="数据文件（JSON，结构含 collected_files: [{stage, source, dest}]）")
    ap.add_argument("--root", required=True,
                    help="本次招标输出根目录；脚本在该目录下创建 01_报名递交资料/ 和 02_投标递交资料/")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不实际复制")
    args = ap.parse_args()

    with open(args.items, "r", encoding="utf-8") as f:
        d = json.load(f)

    items = d.get("collected_files", [])
    if not items:
        print("提示：collected_files 为空，本次无需归集任何合规资料。")
        return

    root = Path(args.root)
    stage_dirs = {
        "报名": root / "01_报名递交资料",
        "投标": root / "02_投标递交资料",
    }
    if not args.dry_run:
        for sd in stage_dirs.values():
            sd.mkdir(parents=True, exist_ok=True)

    skipped = []
    copied = []
    failed = []
    for it in items:
        stage = it.get("stage", "").strip()
        source = it.get("source", "").strip()
        dest = it.get("dest", "").strip()
        if stage not in stage_dirs:
            skipped.append({"item": it, "reason": f"未识别的阶段：{stage}"})
            continue
        if not source:
            failed.append({"item": it, "reason": "源文件路径为空"})
            continue
        src = Path(source)
        if not src.exists():
            failed.append({"item": it, "reason": f"源文件不存在：{source}"})
            continue

        # 决定目标路径
        if dest:
            target = Path(dest)
        else:
            target = stage_dirs[stage] / src.name

        if args.dry_run:
            print(f"[DRY] {stage} {src} -> {target}")
            copied.append({"stage": stage, "source": str(src), "dest": str(target), "dry_run": True})
            continue

        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, target)
            copied.append({"stage": stage, "source": str(src), "dest": str(target)})
            print(f"[OK] {stage} {src} -> {target}")
        except Exception as e:
            failed.append({"item": it, "reason": f"复制失败：{e}"})

    summary = {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "disclaimer": DISCLAIMER,
        "root": str(root),
        "copied_count": len(copied),
        "skipped_count": len(skipped),
        "failed_count": len(failed),
        "copied": copied,
        "skipped": skipped,
        "failed": failed,
    }
    summary_path = root / "归集报告.json"
    if not args.dry_run:
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        with open(summary_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

    print()
    print(f"=== 归集汇总 ===")
    print(f"已复制：{summary['copied_count']} | 跳过：{summary['skipped_count']} | "
          f"失败：{summary['failed_count']}")
    if failed:
        print("失败明细：")
        for f in failed:
            print(f"  - {f['reason']}")
    if not args.dry_run:
        print(f"归集报告写入：{summary_path}")


if __name__ == "__main__":
    main()
