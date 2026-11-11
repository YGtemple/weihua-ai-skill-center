# -*- coding: utf-8 -*-
"""校验《报名资料清单.xlsx》登记的存放路径是否真实存在。

用法：
    python check_ledger.py                      # 使用默认台账
    python check_ledger.py --ledger "D:\\x.xlsx" --sheet 0
输出四类：#OK / #RENAMED（登记名与实存文件不同）/ #MISSING（目录内无对应文件）/ #NOPATH（目录不存在）
"""
import os, sys, argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import readx

DEFAULT_LEDGER = r'E:\卫华供应商入围资料库\招标报名自动整理\报名资料清单.xlsx'


def guess(d, name, base):
    """在目录 d 中按 同名 → 主名一致 → 前缀 的顺序找文件"""
    files = [f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f))]
    keys = [k for k in (base, os.path.splitext(base)[0] if base else '', name) if k]
    for key in keys:
        for f in files:
            if f == key:
                return f
    for key in keys:
        for f in files:
            if os.path.splitext(f)[0] == key:
                return f
    for key in keys:
        for f in files:
            if f.startswith(key):
                return f
    # 去掉括号/空格后再比一次
    norm = lambda s: s.replace(' ', '').replace('　', '')
    for key in keys:
        for f in files:
            if norm(key) in norm(f) or norm(os.path.splitext(f)[0]) in norm(key):
                return f
    return None


def check(ledger_path, sheet_index=0):
    rows = readx.read_xlsx(ledger_path)[sheet_index][1]
    stat = {'OK': 0, 'RENAMED': 0, 'MISSING': 0, 'NOPATH': 0}
    details = []
    for r in rows[1:]:
        if len(r) < 5 or not r[0].strip():
            continue
        name, path = r[0].strip(), r[4].strip()
        if not path:
            stat['NOPATH'] += 1
            details.append(('NOPATH', name, path, ''))
            continue
        if os.path.isfile(path):
            stat['OK'] += 1
            continue
        d = path if os.path.isdir(path) else os.path.dirname(path)
        if not d or not os.path.isdir(d):
            stat['NOPATH'] += 1
            details.append(('NOPATH', name, path, d))
            continue
        hit = guess(d, name, '' if os.path.isdir(path) else os.path.basename(path))
        if hit:
            stat['RENAMED'] += 1
            details.append(('RENAMED', name, path, os.path.join(d, hit)))
        else:
            stat['MISSING'] += 1
            details.append(('MISSING', name, path, d))
    return stat, details


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--ledger', default=DEFAULT_LEDGER)
    ap.add_argument('--sheet', type=int, default=0)
    a = ap.parse_args()
    if not os.path.exists(a.ledger):
        print('台账不存在：%s' % a.ledger)
        return 2
    stat, details = check(a.ledger, a.sheet)
    print('台账：%s' % a.ledger)
    print('正常 %d ｜ 名称不符 %d ｜ 目录内无文件 %d ｜ 路径不存在 %d'
          % (stat['OK'], stat['RENAMED'], stat['MISSING'], stat['NOPATH']))
    for kind in ('RENAMED', 'MISSING', 'NOPATH'):
        group = [x for x in details if x[0] == kind]
        if not group:
            continue
        print('\n[%s]' % kind)
        for _, name, path, extra in group:
            print('  - %s' % name)
            print('      登记：%s' % path)
            print('      实际：%s' % (extra or '（无）'))
    return 0 if stat['RENAMED'] == stat['MISSING'] == stat['NOPATH'] == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
