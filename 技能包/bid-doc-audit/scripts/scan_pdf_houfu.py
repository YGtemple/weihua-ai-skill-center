# -*- coding: utf-8 -*-
"""
PDF "后附"类附件漏插检查（v1.5 升级）。

规则要点（v1.5）：
1. 扫描本次上传报名 PDF 全部可提取文本，检索财务/审计、社保类"XXX：后附"标注；
2. 性能优化：命中后仅向后扫描【后续 5 个页面】寻找附件线索，不读取全部剩余页面；
3. v1.5 局部校验约束：5 页窗口内**只做轻量级特征线索匹配**——不读取报表明细（金额、科目、人员缴费明细）；
4. 附件超出 5 页扫描窗口时，统一标准提示："附件超出扫描窗口，请人工继续翻阅PDF后续页面核对"；
5. AI 不自动判定"缺失"，不生成 ❌ 缺失项，全部线索/告警落入 houfu_scan.json；
6. 三种结果分类输出：✅ 检索到附件线索 / ⚠️ 风险提醒 / 📷 扫描PDF无法识别。

用法：
  python scan_pdf_houfu.py --pdfs 报名01.pdf 报名02.pdf ... --out houfu_scan.json
    --pdfs  一个或多个 PDF 文件路径
    --out   结果 JSON 输出文件路径
    --window  5  后续扫描窗口大小（默认 5，范围 1-10）
"""
import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF (1.24+ 推荐使用 pymupdf)
except ImportError:
    try:
        import fitz  # 旧版 PyMuPDF 兼容
    except ImportError:
        print("错误：缺少 pymupdf 依赖，请先 `pip install pymupdf`", file=sys.stderr)
        sys.exit(2)

# ============ 关键词库 ============
FIN_KEYWORDS = [
    r"近三年财务报表[：:]?后附",
    r"财务报表[：:]?后附",
    r"审计报告[：:]?后附",
    r"\d{2,4}年财务报表[：:]?后附",
    r"\d{2,4}[\-－~到至]\d{2,4}年审计报告[：:]?后附",
]
SOCIAL_KEYWORDS = [
    r"社保清单[：:]?后附",
    r"社保证明[：:]?后附",
    r"社保缴纳记录[：:]?后附",
    r"社保参保证明[：:]?后附",
    r"社保[：:]?后附",
]
ALL_KEYWORDS = [(k, "财务/审计") for k in FIN_KEYWORDS] + [(k, "社保") for k in SOCIAL_KEYWORDS]

# 财务/审计附件线索
FIN_HINTS = ["资产负债表", "利润表", "现金流量表", "审计报告", "会计师事务所",
             "签章", "财务报表", "审计意见", "财务负责人"]
# 社保附件线索
SOCIAL_HINTS = ["社保清单", "社保参保证明", "参保单位", "缴费月份",
                "社保缴费记录", "社保局", "社保缴纳", "参保人员"]

DISCLAIMER = ("本扫描仅输出线索与告警，不做任何 ❌ 缺失判定，所有内容必须人工打开 PDF 复核。"
              "AI仅辅助分析，禁止直接对外提交。")

WINDOW_DEFAULT = 5
WINDOW_MIN, WINDOW_MAX = 1, 10


def classify_keyword(kw):
    if "财务" in kw or "审计" in kw:
        return "财务/审计"
    if "社保" in kw:
        return "社保"
    return "其它"


def get_page_text(page):
    """获取单页文本；空字符串视作扫描型 PDF。"""
    try:
        return page.get_text() or ""
    except Exception:
        return ""


def search_hints(text, hints):
    return [h for h in hints if h in text]


def scan_pdf(pdf_path, window=WINDOW_DEFAULT):
    pdf_path = Path(pdf_path)
    pdf_result = {
        "pdf": str(pdf_path),
        "pdf_name": pdf_path.name,
        "exists": pdf_path.exists(),
        "total_pages": 0,
        "scan_pdf_unrecognized": False,
        "hits": [],
        "error": None,
    }
    if not pdf_path.exists():
        pdf_result["error"] = "PDF 文件不存在"
        return pdf_result

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        pdf_result["error"] = f"PDF 打开失败：{e}"
        return pdf_result

    try:
        total = doc.page_count
        pdf_result["total_pages"] = total

        # 预读全部页文本（仅做命中判定用，不在 hit 上做按页 forward scan）
        page_texts = []
        all_text_combined = ""
        for i in range(total):
            txt = get_page_text(doc.load_page(i))
            page_texts.append(txt)
            all_text_combined += txt + "\n"

        # 判断是否为扫描型 PDF
        if not all_text_combined.strip():
            pdf_result["scan_pdf_unrecognized"] = True
            # 直接退出，给一份统一的兜底备注
            pdf_result["hits"].append({
                "keyword": "（整份PDF文本提取为空）",
                "category": "扫描PDF",
                "page": 0,
                "window": f"全文 N/A",
                "result": "📷 扫描PDF无法识别",
                "evidence": "",
                "note": "【扫描PDF，无法识别\"后附\"文字，请人工自行检查是否插入对应附件】",
                "advice": "人工逐页检查本 PDF 是否插入相应附件（财务报表/审计报告/社保证明）",
            })
            return pdf_result

        # 关键词扫描（财务/审计 + 社保 两组）
        for pattern, category in ALL_KEYWORDS:
            for m in re.finditer(pattern, all_text_combined):
                # 反推页码
                # m.start() 是 all_text_combined 的位置，需要逐页累计位置
                pos = m.start()
                cumulative = 0
                hit_page = 0
                for idx, t in enumerate(page_texts):
                    cumulative += len(t) + 1  # +1 for '\n' we added
                    if pos < cumulative:
                        hit_page = idx + 1
                        break
                if hit_page == 0:
                    continue

                # 5 页窗口扫描（仅此一页向后 5 页）
                window_end = min(hit_page + window, total)
                window_text = "\n".join(page_texts[hit_page:window_end])

                hints = FIN_HINTS if category == "财务/审计" else SOCIAL_HINTS
                matched_hints = search_hints(window_text, hints)

                if matched_hints:
                    pdf_result["hits"].append({
                        "keyword": m.group(0),
                        "category": category,
                        "page": hit_page,
                        "window": f"第{hit_page + 1}页 ~ 第{window_end}页",
                        "result": "✅ 检索到附件线索",
                        "evidence": "|".join(matched_hints),
                        "note": (f"PDF 第{hit_page}页标注\"{m.group(0)}\"，"
                                 f"向后扫描5页范围内已检索到对应附件相关页面，"
                                 f"请人工确认内容是否完整有效。"),
                        "advice": "",
                    })
                else:
                    # v1.5 升级：附件超出 5 页扫描窗口时，统一使用标准提示文案
                    if window_end < total:
                        advice = (f"附件超出扫描窗口，请人工继续翻阅PDF后续页面（第{window_end + 1}页及之后）核对")
                        note = (f"⚠️【附件超出扫描窗口】PDF 第{hit_page}页标注\"{m.group(0)}\"，"
                                f"附件可能放置在第{window_end + 1}页及之后，"
                                f"请人工继续翻阅PDF后续页面核对。")
                    else:
                        advice = "已扫描PDF全部页码仍未命中，请人工手动逐页检查是否插入对应附件"
                        note = (f"⚠️【风险提醒】PDF 第{hit_page}页写有\"{m.group(0)}\"，"
                                f"向后连续5页未检索到对应附件页面，"
                                f"极有可能忘记插入对应附件；若附件放在更靠后的位置，请人工手动继续往后翻PDF核查！")
                    pdf_result["hits"].append({
                        "keyword": m.group(0),
                        "category": category,
                        "page": hit_page,
                        "window": f"第{hit_page + 1}页 ~ 第{window_end}页",
                        "result": "⚠️ 风险提醒",
                        "evidence": "",
                        "note": note,
                        "advice": advice,
                    })

    finally:
        doc.close()

    return pdf_result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdfs", nargs="+", required=True, help="一份或多份 PDF 文件路径")
    ap.add_argument("--out", default=None, help="结果 JSON 输出路径")
    ap.add_argument("--window", type=int, default=WINDOW_DEFAULT,
                    help=f"向后扫描窗口大小（默认 {WINDOW_DEFAULT}，范围 {WINDOW_MIN}-{WINDOW_MAX}）")
    args = ap.parse_args()

    window = max(WINDOW_MIN, min(args.window, WINDOW_MAX))

    result = {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "window_pages": window,
        "disclaimer": DISCLAIMER,
        "results": [],
    }

    for pdf in args.pdfs:
        r = scan_pdf(pdf, window=window)
        result["results"].append(r)

    # 统计
    total_hits = sum(len(r["hits"]) for r in result["results"])
    risk_count = sum(1 for r in result["results"] for h in r["hits"] if h["result"].startswith("⚠️"))
    found_count = sum(1 for r in result["results"] for h in r["hits"] if h["result"].startswith("✅"))
    scan_count = sum(1 for r in result["results"] for h in r["hits"] if h["result"].startswith("📷"))
    result["stats"] = {
        "pdfs_scanned": len(result["results"]),
        "total_hits": total_hits,
        "found": found_count,
        "risk": risk_count,
        "scan_unrecognized": scan_count,
    }

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"已写入 {args.out}")
    print(f"[scan_pdf_houfu] 扫描 PDF 数={len(args.pdfs)}, 命中关键词={total_hits} 次 "
          f"(✅{found_count} 检索到附件线索 / ⚠️{risk_count} 风险提醒 / 📷{scan_count} 扫描PDF)")
    print(text)


if __name__ == "__main__":
    main()
