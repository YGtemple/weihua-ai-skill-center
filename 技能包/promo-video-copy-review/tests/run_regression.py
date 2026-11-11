#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""可重复运行的回归测试。

用法（在技能根目录或任意目录均可）：
    python tests/run_regression.py

只做黑盒校验，绝不向真实 ledger/审核台账.csv 写任何东西。
校验项：
  1. check_copy.py 对 S001.txt（27.4s）命中 2 处
  2. check_copy.py 对 S002.txt（28.1s）命中 0 处
  3. check_copy.py 对 negative.txt 命中 8 处
  4. log_review.py --stats-only 对两篇快照：累计2篇、平均58.0、P0共6、核片率100%
  5. 运行前后真实 ledger/审核台账.csv 的 SHA256 不变
退出码：全部通过=0，任一失败=1。
"""

import hashlib
import os
import subprocess
import sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(SKILL_DIR, "scripts")
FIX = os.path.join(SKILL_DIR, "tests", "fixtures")
REAL_LEDGER = os.path.join(SKILL_DIR, "ledger", "审核台账.csv")


def run(args):
    proc = subprocess.run(
        [sys.executable] + args,
        cwd=SKILL_DIR, capture_output=True, text=True, encoding="utf-8",
    )
    return proc.returncode, proc.stdout + proc.stderr


def sha256(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


results = []


def check(name, cond, detail=""):
    results.append((name, cond, detail))
    print(("PASS  " if cond else "FAIL  ") + name + (("  | " + detail) if detail else ""))


def hit_count(out):
    for line in out.splitlines():
        if "风险词命中" in line and "处" in line:
            return line.strip()
    return ""


def main():
    ledger_before = sha256(REAL_LEDGER) if os.path.exists(REAL_LEDGER) else "(无真实台账)"

    # 1-3: check_copy 三个样例
    rc, out1 = run([os.path.join(SCRIPTS, "check_copy.py"),
                    "--file", os.path.join(FIX, "S001.txt"), "--duration", "27.4"])
    check("S001 命中 2 处", "风险词命中 2 处" in out1, hit_count(out1))

    rc, out2 = run([os.path.join(SCRIPTS, "check_copy.py"),
                    "--file", os.path.join(FIX, "S002.txt"), "--duration", "28.1"])
    check("S002 命中 0 处", "风险词命中 0 处" in out2, hit_count(out2))

    rc, out3 = run([os.path.join(SCRIPTS, "check_copy.py"),
                    "--file", os.path.join(FIX, "negative.txt")])
    check("negative 命中 8 处", "风险词命中 8 处" in out3, hit_count(out3))

    # 4: log_review 统计（用快照，不碰真实台账）
    rc, out4 = run([os.path.join(SCRIPTS, "log_review.py"),
                    "--ledger", os.path.join(FIX, "ledger-fixture.csv"),
                    "--stats-only"])
    check("统计=累计2篇", "累计审稿 2 篇" in out4,
          [l for l in out4.splitlines() if "累计" in l][0] if "累计" in out4 else out4)
    check("统计=平均58.0", "平均得分 58.0" in out4)
    check("统计=P0共6项", "拦下 P0 共 6 项" in out4)
    check("统计=核片率100%", "核片率 2/2（100%）" in out4)

    # 5: 真实台账未被改动
    ledger_after = sha256(REAL_LEDGER) if os.path.exists(REAL_LEDGER) else "(无真实台账)"
    check("真实台账前后一致", ledger_before == ledger_after,
          "sha256=" + ledger_after[:16] + "…")

    failed = [n for n, ok, _ in results if not ok]
    print("\n========================================")
    print(f"回归结果：{len(results) - len(failed)}/{len(results)} 通过")
    if failed:
        print("未通过：" + "；".join(failed))
        return 1
    print("全部通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
