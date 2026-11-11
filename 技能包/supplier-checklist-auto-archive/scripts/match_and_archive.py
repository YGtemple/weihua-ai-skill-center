# -*- coding: utf-8 -*-
"""通用版-供应商资料清单自动归档：Step 2 台账匹配 + 文件归集（v2 四类匹配）。

依据：
  1. E:\卫华供应商入围资料库\供应商入围资料清单.xlsx（佐证资料台账：文件名/文件完整路径/备注/甲方常用别名(可选)）
  2. E:\卫华供应商入围资料库\资质清单.xlsx（资质证书台账：类别/证书完整名称/文件路径/状态/备注(可选)/甲方常用别名(可选)）

规则：
  - 匹配仅以上述两份台账为唯一依据，禁止编造文件；
  - 匹配结果分 4 类：
      ✅ 已匹配：名称精准命中；
      ⚠️ 语义模糊匹配：语义近似但无法100%确认，需人工复核；
      📎 台账字段候选文件：台账「备注」「甲方常用别名」字段含相似字眼的备选文件，
         仅归集供人工甄别，不判定为符合要求；
      ❌ 无匹配：台账内无对应文件；
  - 仅复制副本到输出目录（默认桌面\本次资料归集输出），绝不移动/修改/删除源文件；
  - 副本重命名为「序号-原文件名」；📎 候选文件加前缀「序号-【候选-人工甄别】原文件名」；
  - 生成《资料归集完成汇总表.xlsx》，首行强制 AI 复核提示；📎 行明确标记来源字段。

默认归集 ✅ 项与 📎 候选项；⚠️ 模糊匹配项需人工确认后用 --confirm-fuzzy 传入序号列表再归档。
"""
import argparse
import json
import os
import re
import shutil
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

DEFAULT_BASE = r"E:\卫华供应商入围资料库"
LEDGER_QUAL = "资质清单.xlsx"
LEDGER_MAT = "供应商入围资料清单.xlsx"
FUZZY_THRESHOLD = 0.5  # 关键词级相似度（LCS覆盖率/双字元Dice），≥此值列为⚠️候选
NOTE_FUZZY_THRESHOLD = 0.45  # 📎 台账备注/别名候选阈值：候选仅供人工甄别，宁多勿漏，略低于⚠️
CANDIDATE_PREFIX = "【候选-人工甄别】"
AI_NOTICE = ("【本结果由AI辅助生成，包含部分来自台账备注、别名的候选备选文件，候选文件不一定满足甲方要求，"
             "务必人工逐条甄别筛选，剔除无关文件，补充缺失资料后，才可对外提交，不可直接使用归集出来的全套文件】")


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("台帐", "台账").replace("帐", "账")
    return re.sub(r"[\s\u3000]+", "", s)


def sanitize_name(s):
    return re.sub(r'[\\/:*?"<>|]+', "_", str(s)).strip() or "未命名"


def get_desktop():
    home = os.path.expanduser("~")
    for cand in (os.path.join(home, "Desktop"), os.path.join(home, "桌面")):
        if os.path.isdir(cand):
            return cand
    return home


# ---------- 台账加载 ----------

def read_ledger(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise SystemExit(f"[告警] 台账为空：{path}")
    header = [norm(h) for h in rows[0]]
    data = []
    for row in rows[1:]:
        if not any(cell is not None and str(cell).strip() for cell in row):
            continue
        data.append({header[i]: ("" if i >= len(row) or row[i] is None else str(row[i]).strip())
                     for i in range(len(header))})
    return header, data


def find_col(header, keys):
    for i, h in enumerate(header):
        if any(k in h for k in keys):
            return i
    return None


def load_entries(base):
    """返回 entry 列表：
      name      主匹配名（证书名称/文件名去扩展名）
      alias     一般匹配语料（类别等，参与 ✅/⚠️ 匹配）
      extra     候选字段 [(text, 来源标签)]，来自「备注」「甲方常用别名」，仅参与 📎 候选
      ...
    """
    entries = []

    pq = os.path.join(base, LEDGER_QUAL)
    if os.path.isfile(pq):
        header, data = read_ledger(pq)
        cols = {k: find_col(header, keys) for k, keys in {
            "cat": ["类别"], "name": ["证书完整名称", "证书名称", "资质名称"],
            "path": ["文件路径"], "status": ["状态"], "note": ["备注"], "alias": ["甲方常用别名", "别名"],
        }.items()}
        for row in data:
            name = row.get(header[cols["name"]], "") if cols["name"] is not None else ""
            if not name:
                continue
            alias, extra = [], []
            if cols["cat"] is not None and row.get(header[cols["cat"]]):
                alias.append(row[header[cols["cat"]]])
            if cols["alias"] is not None and row.get(header[cols["alias"]]):
                extra.append((row[header[cols["alias"]]], "甲方常用别名"))
            if cols["note"] is not None and row.get(header[cols["note"]]):
                extra.append((row[header[cols["note"]]], "台账备注"))
            entries.append({
                "name": name,
                "alias": [a for a in alias if a],
                "extra": extra,
                "file_rel": row.get(header[cols["path"]], "") if cols["path"] is not None else "",
                "ledger": "资质清单",
                "status": row.get(header[cols["status"]], "") if cols["status"] is not None else "",
            })
    else:
        print(f"[警告] 未找到台账 {pq}")

    pm = os.path.join(base, LEDGER_MAT)
    if os.path.isfile(pm):
        header, data = read_ledger(pm)
        cols = {k: find_col(header, keys) for k, keys in {
            "name": ["文件名"], "path": ["文件完整路径", "路径"], "note": ["备注"], "alias": ["甲方常用别名", "别名"],
        }.items()}
        for row in data:
            name = row.get(header[cols["name"]], "") if cols["name"] is not None else ""
            if not name:
                continue
            extra = []
            if cols["alias"] is not None and row.get(header[cols["alias"]]):
                extra.append((row[header[cols["alias"]]], "甲方常用别名"))
            if cols["note"] is not None and row.get(header[cols["note"]]):
                extra.append((row[header[cols["note"]]], "台账备注"))
            entries.append({
                "name": os.path.splitext(name)[0],
                "alias": [],
                "extra": extra,
                "file_rel": "",
                "file_abs_hint": row.get(header[cols["path"]], "") if cols["path"] is not None else "",
                "file_name_raw": name,
                "ledger": "资料清单",
            })
    else:
        print(f"[警告] 未找到台账 {pm}")

    if not entries:
        raise SystemExit("[告警] 两份台账均无法加载有效条目，停止执行。")
    return entries


# ---------- 文件定位（只读） ----------

class FileLocator:
    """台账路径修复：资料清单路径存在拼接错误时，按文件名在资料库递归搜索。"""

    def __init__(self, base):
        self.base = base
        self._index = None
        self.fixed = 0

    def _build_index(self):
        idx = {}
        for root, _dirs, files in os.walk(self.base):
            for fn in files:
                idx.setdefault(fn.lower(), os.path.join(root, fn))
        self._index = idx

    def resolve(self, entry):
        # 资质清单：相对路径
        if entry.get("file_rel"):
            p = os.path.join(self.base, entry["file_rel"])
            if os.path.isfile(p):
                return p, False
        raw = entry.get("file_name_raw") or os.path.basename(entry.get("file_rel", "")) or entry["name"]
        # 资料清单：台账绝对路径（可能拼接错误）
        hint = entry.get("file_abs_hint", "")
        if hint:
            for cand in (hint, os.path.join(hint, raw)):
                cand = cand.rstrip("\\") if os.path.isdir(cand) else cand
                if os.path.isfile(cand):
                    return cand, False
            # 修复：文件名+扩展名在 hint 目录附近搜索
            if os.path.isfile(os.path.join(hint, raw)):
                return os.path.join(hint, raw), False
        # 兜底：全库递归按文件名搜索
        if self._index is None:
            self._build_index()
        for fn in (raw, f"{entry['name']}.pdf", f"{entry['name']}.jpg", f"{entry['name']}.png"):
            hit = self._index.get(fn.lower())
            if hit:
                self.fixed += 1
                return hit, True
        return None, False


# ---------- 匹配 ----------

# 要求列常见的动作动词前缀（如「查看合格供方名单」「现场查看生产车间」），
# 剥离后得到的资料关键词才是与台账名称比对的核心。
ACTION_PREFIXES = ("现场查看", "现场检查", "现场核查", "查阅", "查看", "核对", "查验",
                   "检查", "提供", "索要", "抽查", "询问", "核实")
# 过于泛化的关键词（剥离动词后只剩这些词时，禁止模糊匹配与候选挂靠，
# 避免把「相关证书」这类泛词挂到任意证书上）
GENERIC_KW = {"相关证书", "证书", "文件", "记录", "资料", "数据", "情况", "制度", "证明文件", "清单"}


def strip_action(req_n):
    for p in sorted(ACTION_PREFIXES, key=len, reverse=True):
        if req_n.startswith(p):
            return req_n[len(p):]
    return req_n


def split_fragments(kw):
    """复合要求（如「设备管理制度，抽查监控记录、检维修记录」）按标点拆成片段，
    每个片段再剥离动词前缀、去掉「相关」等虚词与「2分：」类评分标记。"""
    frags = []
    for seg in re.split(r"[，,、；;。/（）()：:]+", kw):
        seg = re.sub(r"^[0-9.+\-]+分[：:]?", "", seg)  # 剥离「2分：」「1.5分:」等评分标记
        seg = strip_action(seg).replace("相关", "")
        seg = seg.strip("：: .．0-9")
        if len(seg) >= 4:
            frags.append(seg)
    return frags


def bigrams(s):
    return {s[i:i + 2] for i in range(len(s) - 1)} if len(s) > 1 else {s}


def score(kw, t):
    """关键词 kw 与台账文本 t 的相似度：最长公共子串覆盖率 与 双字元 Dice 的较大值。
    公共子串须 ≥3 字（或覆盖整个关键词）才计入，避免「证书」「情况」两字重合的噪音。"""
    if not kw or not t:
        return 0.0
    m = [[0] * (len(t) + 1) for _ in range(len(kw) + 1)]
    best = 0
    for i in range(1, len(kw) + 1):
        for j in range(1, len(t) + 1):
            if kw[i - 1] == t[j - 1]:
                m[i][j] = m[i - 1][j - 1] + 1
                if m[i][j] > best:
                    best = m[i][j]
    lcs_ratio = (best / len(kw)) if (best >= 3 or best == len(kw)) else 0.0
    b1, b2 = bigrams(kw), bigrams(t)
    dice = 2 * len(b1 & b2) / (len(b1) + len(b2)) if b1 and b2 else 0.0
    return max(lcs_ratio, dice)


def match_one(req_norm, entries):
    """返回 (status, best, cands)。
    status: exact / exact_multi / fuzzy / cand / none
    - exact|exact_multi: best=命中条目，cands=其余候选条目列表
    - fuzzy: best=最相似条目，cands=[(条目, 相似度), ...]
    - cand: best=None，cands=[(条目, 来源标签, 命中字段原文), ...]（📎台账字段候选）"""
    kw = strip_action(req_norm).replace("相关", "")
    frags = split_fragments(req_norm) or ([kw] if kw else [])
    if not frags:
        return "none", None, []
    exact_hits, fuzzy_hits = [], []
    generic = all(f in GENERIC_KW for f in frags) if frags else True
    for e in entries:
        texts = [norm(e["name"])] + [norm(a) for a in e["alias"] if norm(a)]
        best_r, best_t = 0.0, ""
        for t in texts:
            if not t:
                continue
            hit = False
            # 精准：片段（≥4字）包含于台账名，或台账名（≥4字）包含于完整要求
            for f in frags:
                if len(f) >= 4 and f in t:
                    exact_hits.append((e, t))
                    hit = True
                    break
            if not hit and len(t) >= 4 and t in kw:
                exact_hits.append((e, t))
                hit = True
            if hit:
                continue
            r = max(score(f, t) for f in frags)
            if r > best_r:
                best_r, best_t = r, t
        if not any(h[0] is e for h in exact_hits) and not generic and best_r >= FUZZY_THRESHOLD:
            fuzzy_hits.append((e, best_r, best_t))
    if exact_hits:
        # 去重（同一文件多条记录），按命中文本长度优先
        exact_hits.sort(key=lambda h: -len(h[1]))
        uniq, seen = [], set()
        for e, t in exact_hits:
            key = (e["ledger"], e["name"])
            if key not in seen:
                seen.add(key)
                uniq.append(e)
        if len(uniq) == 1:
            return "exact", uniq[0], []
        return "exact_multi", uniq[0], uniq[1:]
    if fuzzy_hits:
        fuzzy_hits.sort(key=lambda h: -h[1])
        top = [h for h in fuzzy_hits if h[1] >= fuzzy_hits[0][1] - 0.05][:3]
        return "fuzzy", top[0][0], [(h[0], round(h[1], 2)) for h in top[1:]]
    # 📎 台账字段候选：备注/甲方常用别名中含相似字眼（片段包含或高相似），供人工甄别
    note_cands = find_note_candidates(frags, generic, entries)
    if note_cands:
        return "cand", None, note_cands
    return "none", None, []


def find_note_candidates(frags, generic, entries):
    """在台账「备注」「甲方常用别名」字段中检索相似字眼，返回 [(entry, 来源标签, 字段原文)]。
    仅作候选备选，绝不判定为符合要求；泛词（如「证书」）不触发候选，防噪音。"""
    cands, seen = [], set()
    for e in entries:
        for text, label in e["extra"]:
            t = norm(text)
            if not t:
                continue
            hit = False
            for f in frags:
                if len(f) >= 3 and f not in GENERIC_KW and (f in t or t in f):
                    hit = True
                    break
            if not hit and not generic and frags and max(score(f, t) for f in frags) >= NOTE_FUZZY_THRESHOLD:
                hit = True
            if hit:
                key = (e["ledger"], e["name"], label)
                if key not in seen:
                    seen.add(key)
                    cands.append((e, label, text))
    return cands


# ---------- 主流程 ----------

def copy_file(src, cat_dir, new_name):
    os.makedirs(cat_dir, exist_ok=True)
    dst = os.path.join(cat_dir, sanitize_name(new_name))
    if os.path.isfile(dst):
        base_name, ext = os.path.splitext(dst)
        i = 2
        while os.path.isfile(f"{base_name}({i}){ext}"):
            i += 1
        dst = f"{base_name}({i}){ext}"
    shutil.copy2(src, dst)  # 只复制副本，源文件不动
    return dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--items", required=True, help="Step 1 输出的 parsed_items.json")
    ap.add_argument("--out", default=None, help="输出根目录（默认 桌面/本次资料归集输出）")
    ap.add_argument("--base", default=DEFAULT_BASE, help="资料库根目录")
    ap.add_argument("--confirm-fuzzy", default="", help="人工确认归档的模糊匹配序号，逗号分隔，如 1.2,2.3")
    ap.add_argument("--auto-fuzzy", action="store_true",
                    help="用户明确跳过人工确认时：全部模糊项自动归档，但带【候选-人工甄别】前缀标记")
    ap.add_argument("--evidence", action="store_true",
                    help="佐证扩展模式：对❌无匹配项降低相似度阈值(0.30)捞佐证候选文件，"
                         "全部带【候选-人工甄别】前缀，标注低置信度，必须人工甄别")
    args = ap.parse_args()

    if args.evidence:
        global FUZZY_THRESHOLD
        FUZZY_THRESHOLD = 0.30

    with open(args.items, encoding="utf-8") as f:
        parsed = json.load(f)
    items = parsed["items"]
    if not items:
        raise SystemExit("[告警] parsed_items.json 中没有条目")

    confirmed = {s.strip() for s in args.confirm_fuzzy.split(",") if s.strip()}
    out_root = args.out or os.path.join(get_desktop(), "本次资料归集输出")
    os.makedirs(out_root, exist_ok=True)

    entries = load_entries(args.base)
    n_extra = sum(len(e["extra"]) for e in entries)
    if n_extra == 0:
        print("[提示] 两份台账的「备注」「甲方常用别名」字段当前均为空/不存在，📎 台账字段候选将不会有结果；"
              "在台账中补充这些字段后自动生效。")
    locator = FileLocator(args.base)

    results = []  # 每项含 files: [{src, dst, source_label}]
    for it in items:
        req_n = norm(it["requirement"])
        status, best, cands = match_one(req_n, entries)
        row = {"seq": it["seq"], "category": it["category"], "requirement": it["requirement"],
               "mark": "", "detail": "", "files": []}
        if status in ("exact", "exact_multi"):
            src, _ = locator.resolve(best)
            if src is None:
                row["mark"] = "❌"
                row["detail"] = f"台账《{best['ledger']}》有记录「{best['name']}」，但磁盘上找不到该文件【资料库缺失，需手动补充】"
            else:
                row["mark"] = "✅"
                row["detail"] = f"已匹配《{best['ledger']}》：「{best['name']}」" + (
                    f"；另有 {len(cands)} 个候选：{'、'.join(c['name'] for c in cands)}，请复核" if status == "exact_multi" else "")
                row["files"].append({"src": src, "label": ""})
        elif status == "fuzzy":
            row["mark"] = "⚠️"
            cand_names = [best["name"]] + [c[0]["name"] for c in cands]
            row["detail"] = f"语义模糊匹配候选：{'、'.join(cand_names)}（需人工复核确认）"
            src, _ = locator.resolve(best)
            if src:
                row["src"] = src
            if it["seq"] in confirmed:
                row["detail"] += "；已人工确认归档"
                if src:
                    row["files"].append({"src": src, "label": ""})
            elif args.auto_fuzzy and src:
                # 用户明确跳过人工确认环节：模糊项自动归档，但必须带候选标记，仍需人工甄别
                row["detail"] += "；按指令跳过人工确认自动归档为候选，务必人工甄别"
                if args.evidence:
                    row["detail"] += "（低置信佐证候选）"
                row["files"].append({"src": src, "label": "语义模糊匹配"})
        elif status == "cand":
            row["mark"] = "📎"
            parts = []
            for e, label, text in cands:
                src, _ = locator.resolve(e)
                if src is None:
                    parts.append(f"「{e['name']}」（{label}：「{text}」）—— 磁盘上找不到该文件，需手动补充")
                else:
                    parts.append(f"「{e['name']}」（{label}：「{text}」）")
                    row["files"].append({"src": src, "label": label, "note_text": text})
            row["detail"] = "台账字段候选文件（来自备注/别名，必须人工甄别是否为本条所需）：" + "；".join(parts)
        else:
            row["mark"] = "❌"
            row["detail"] = "资料库缺失，需手动补充"
        results.append(row)

    # 文件归集（仅复制副本）
    copied = 0
    for row in results:
        if not row["files"]:
            continue
        cat_dir = os.path.join(out_root, sanitize_name(row["category"]))
        for fi in row["files"]:
            base = os.path.basename(fi["src"])
            if fi["label"]:  # 📎 候选文件加前缀
                new_name = f"{row['seq']}-{CANDIDATE_PREFIX}{base}"
            else:
                new_name = f"{row['seq']}-{base}"
            dst = copy_file(fi["src"], cat_dir, new_name)
            fi["dst"] = dst
            copied += 1

    # 汇总表（📎 多候选文件时每个文件一行，序号/分类/要求重复以便筛选）
    xlsx_path = os.path.join(out_root, "资料归集完成汇总表.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总表"
    warn_fill = PatternFill("solid", fgColor="FDE9D9")
    cand_fill = PatternFill("solid", fgColor="E4DFEC")
    ws.merge_cells("A1:G1")
    c = ws.cell(1, 1, AI_NOTICE)
    c.font = Font(bold=True, color="C00000")
    c.fill = warn_fill
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    headers = ["序号", "分类名称", "要求内容", "匹配结果", "文件原路径", "输出文件夹路径", "人工复核确认栏"]
    ws.append(headers)
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
    warn_font = Font(color="BF8F00", bold=True)
    bad_font = Font(color="C00000")
    cand_font = Font(color="7030A0", bold=True)

    def file_rows(row):
        """按文件展开；无文件的行也输出一行（路径留空）。"""
        if not row["files"]:
            return [{"mark": row["mark"], "detail": row["detail"], "src": "", "outdir": ""}]
        out = []
        for fi in row["files"]:
            if fi["label"]:
                detail = f"【候选-人工甄别｜来源:{fi['label']}：「{fi.get('note_text', '')}」】"
            else:
                detail = row["detail"]
            out.append({"mark": row["mark"], "detail": detail,
                        "src": fi["src"], "outdir": os.path.dirname(fi["dst"])})
        return out

    for row in results:
        for fr in file_rows(row):
            ws.append([row["seq"], row["category"], row["requirement"],
                       f"{fr['mark']} {fr['detail']}".strip(), fr["src"], fr["outdir"] or "", ""])
            r = ws.max_row
            if row["mark"] == "⚠️":
                ws.cell(r, 4).font = warn_font
            elif row["mark"] == "❌":
                ws.cell(r, 4).font = bad_font
            elif row["mark"] == "📎":
                ws.cell(r, 4).font = cand_font
                for col in range(1, 8):
                    ws.cell(r, col).fill = cand_fill
    for col, w in zip(range(1, 8), (10, 24, 46, 50, 46, 40, 16)):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A3"
    wb.save(xlsx_path)

    n_ok = sum(1 for r in results if r["mark"] == "✅")
    n_fz = sum(1 for r in results if r["mark"] == "⚠️")
    n_cd = sum(1 for r in results if r["mark"] == "📎")
    n_no = sum(1 for r in results if r["mark"] == "❌")
    print(f"[归集完成] 输出目录：{out_root}")
    print(f"  ✅ 已匹配：{n_ok} 项 | ⚠️ 模糊匹配：{n_fz} 项 | 📎 台账字段候选：{n_cd} 项 | ❌ 无匹配：{n_no} 项")
    print(f"  已复制副本文件 {copied} 个（含📎候选备选文件；源文件零改动）")
    if locator.fixed:
        print(f"  台账路径修复：{locator.fixed} 个文件按文件名在资料库内自动定位")
    if n_fz:
        pending = [r["seq"] for r in results if r["mark"] == "⚠️" and r["seq"] not in confirmed
                   and not (args.auto_fuzzy and any(f.get("label") == "语义模糊匹配" for f in r["files"]))]
        if pending:
            print(f"  [待确认] 模糊匹配序号：{','.join(pending)} —— 人工确认后可用 --confirm-fuzzy {','.join(pending)} 重新执行归档")
    print(f"  汇总表：{xlsx_path}")
    print("  " + AI_NOTICE)


if __name__ == "__main__":
    main()
