# -*- coding: utf-8 -*-
"""甲方入围需求匹配 + 附件副本收集。

用法:
    python match_requirement.py --out <输出目录> \
        --req <需求清单文件(.txt 每行一条 / .json 数组)> \
        [--ledgers <parse_ledgers.py 的输出目录>] \
        [--base E:\\卫华供应商入围资料库] \
        [--collect-dir <附件副本目录>] \
        [--rename-pattern "1-{原名}"]      # 占位符: {原名} {无扩展名} {序号} {扩展名}

匹配标记: ✅已匹配 / ⚠️需人工补充(多候选或需人工确认) / ❌资料库缺失
硬性规则: 匹配不确定即标记缺失/人工补充，禁止编造；只读源文件，输出全部为副本。
"""
import argparse
import json
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import DEFAULT_BASE, MARK_MANUAL, MARK_MISSING, MARK_OK, die, ensure_dir  # noqa: E402

# 需求文本 → 台账条目的匹配：双向包含（去除空白与常见噪声词）
NOISE = re.compile(r"[，。、；：,.;:（）()\s　「」『』“”\"'【】\[\]一二三四五六七八九十0-9]+")


def clean(s):
    return NOISE.sub("", str(s))


def match_item(req_clean, candidates, key_fields):
    """返回候选列表: [(rec, 命中字段), ...]"""
    hits = []
    for rec in candidates:
        for kf in key_fields:
            val = rec.get(kf) or ""
            vc = clean(val)
            if vc and (req_clean in vc or vc in req_clean):
                hits.append((rec, kf))
                break
    return hits


def load_requirements(path):
    if not os.path.isfile(path):
        die(f"需求清单文件不存在: {path}")
    text = open(path, encoding="utf-8-sig").read().strip()
    if path.lower().endswith(".json"):
        data = json.loads(text)
        if isinstance(data, list):
            return [str(x) for x in data if str(x).strip()]
        if isinstance(data, dict):
            return [str(v) for v in data.values() if str(v).strip()]
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--req", required=True, help="需求清单文件")
    ap.add_argument("--ledgers", default=None, help="parse_ledgers.py 的输出目录（默认 --out）")
    ap.add_argument("--base", default=DEFAULT_BASE)
    ap.add_argument("--collect-dir", default=None, help="附件副本收集目录（默认 <out>/附件副本）")
    ap.add_argument("--rename-pattern", default="{原名}", help="重命名模板，如 '1-{原名}'")
    args = ap.parse_args()

    led_dir = args.ledgers or args.out
    for name in ("ledgers_company.json", "ledgers_qualifications.json", "ledgers_materials.json"):
        p = os.path.join(led_dir, name)
        if not os.path.isfile(p):
            die(f"缺少台账解析产物: {p}。请先运行 parse_ledgers.py --out {led_dir}")

    company = json.load(open(os.path.join(led_dir, "ledgers_company.json"), encoding="utf-8"))
    quals = json.load(open(os.path.join(led_dir, "ledgers_qualifications.json"), encoding="utf-8"))["qualifications"]
    mats = json.load(open(os.path.join(led_dir, "ledgers_materials.json"), encoding="utf-8"))["materials"]

    out = ensure_dir(args.out)
    collect = ensure_dir(args.collect_dir or os.path.join(out, "附件副本"))
    reqs = load_requirements(args.req)

    results = []
    used_names = {}
    for idx, raw in enumerate(reqs, start=1):
        rq = clean(raw)
        rec = {"序号": idx, "甲方需求": raw, "标记": "", "说明": "", "附件源路径": "", "附件副本": ""}

        # 匹配优先级: 资质 > 资料 > 公司信息（甲方"安全生产许可证"应命中资质而非公司字段）
        q_hits = match_item(rq, quals, ["类别", "证书完整名称"])
        m_hits = match_item(rq, mats, ["文件名"])
        comp_hits = [(k, v) for k, v in company["company_info"].items() if rq and (rq in clean(k) or clean(k) in rq)]

        if len(q_hits) == 1:
            qrec, kf = q_hits[0]
            rec["标记"] = MARK_OK
            rec["说明"] = f"资质匹配: {qrec.get('类别','')} | {qrec.get('证书完整名称','')} (命中列: {kf})"
            rec["附件源路径"] = qrec.get("_resolved_path") or ""
            rec["资质编号"] = qrec.get("证书编号", "")
            rec["有效期"] = qrec.get("有效期", "")
            if qrec.get("_resolved_path"):
                rec["附件副本"] = collect_copy(collect, qrec["_resolved_path"], raw, args.rename_pattern, idx, used_names)
        elif len(q_hits) > 1:
            rec["标记"] = MARK_MANUAL
            cands = [f"{r.get('类别','')}/{r.get('证书完整名称','')}" for r, _ in q_hits[:5]]
            rec["说明"] = f"资质多候选，需人工确认: " + "；".join(cands)
        elif len(m_hits) == 1:
            mrec, kf = m_hits[0]
            rec["标记"] = MARK_OK
            rec["说明"] = f"资料匹配: {mrec.get('文件名','')} (命中列: {kf})"
            rec["附件源路径"] = mrec.get("_resolved_path") or ""
            if mrec.get("_resolved_path"):
                rec["附件副本"] = collect_copy(collect, mrec["_resolved_path"], raw, args.rename_pattern, idx, used_names)
        elif len(m_hits) > 1:
            rec["标记"] = MARK_MANUAL
            cands = [r.get("文件名", "") for r, _ in m_hits[:5]]
            rec["说明"] = f"资料多候选，需人工确认: " + "；".join(cands)
        elif comp_hits:
            rec["标记"] = MARK_OK
            k, v = comp_hits[0]
            rec["说明"] = f"公司信息字段『{k}』= {v}"
            rec["公司字段"] = k
            rec["字段值"] = v
        else:
            rec["标记"] = MARK_MISSING
            rec["说明"] = "资料库无匹配项，需人工补充（禁止编造）"

        results.append(rec)

    # 报告
    lines = ["# 甲方入围需求匹配报告", ""]
    lines.append(f"- 需求条目数: {len(results)}")
    ok = sum(1 for r in results if r["标记"] == MARK_OK)
    man = sum(1 for r in results if r["标记"] == MARK_MANUAL)
    mis = sum(1 for r in results if r["标记"] == MARK_MISSING)
    lines.append(f"- {MARK_OK}: {ok} | {MARK_MANUAL}: {man} | {MARK_MISSING}: {mis}")
    lines.append("")
    lines.append("| 序号 | 甲方需求 | 标记 | 说明 | 附件副本 |")
    lines.append("|---|---|---|---|---|")
    for r in results:
        lines.append(f"| {r['序号']} | {r['甲方需求']} | {r['标记']} | {r['说明']} | {r['附件副本']} |")
    lines.append("")
    lines.append("> 【本结果由AI生成，需人工复核、补充缺失资料，不做最终定稿，不可直接对外提交甲方】")

    with open(os.path.join(out, "match_report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    with open(os.path.join(out, "match_result.json"), "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"✅ 匹配完成: {MARK_OK} {ok} | {MARK_MANUAL} {man} | {MARK_MISSING} {mis}")
    print(f"   报告: {os.path.join(out, 'match_report.md')}")
    print(f"   附件副本: {collect}（{len(os.listdir(collect)) if os.path.isdir(collect) else 0} 个文件）")
    if mis:
        print("⚠️ 存在缺失项，请人工补充，禁止编造。")


def collect_copy(collect, src, req_raw, pattern, idx, used_names):
    """复制附件副本并按重命名模板改名；返回新文件名（空字符串表示失败）。"""
    if not src or not os.path.isfile(src):
        return ""
    ext = os.path.splitext(src)[1]
    base = os.path.splitext(os.path.basename(src))[0]
    try:
        new_name = pattern.format(原名=os.path.basename(src), 无扩展名=base, 序号=idx, 扩展名=ext.lstrip("."))
    except KeyError as e:
        die(f"--rename-pattern 含未知占位符: {e}；支持 {{原名}} {{无扩展名}} {{序号}} {{扩展名}}")
    new_name = sanitize(new_name)
    # 重名自动加序号
    n = used_names.get(new_name, 0) + 1
    used_names[new_name] = n
    if n > 1:
        stem, e2 = os.path.splitext(new_name)
        new_name = f"{stem}({n}){e2}"
    dest = os.path.join(collect, new_name)
    try:
        shutil.copy2(src, dest)
    except Exception as e:
        die(f"附件复制失败 {src} → {dest}: {e}")
    return new_name


def sanitize(name):
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name.strip() or "未命名"


if __name__ == "__main__":
    main()
