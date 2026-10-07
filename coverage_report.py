#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成 FAR Part 23 修订案 → NPRM / Final Rule / 更正件 的覆盖报告。

匹配用两条线取并集，避免单一来源漏判：
  A. 案号线：FAR23_条款-规则文件对应表.csv 里的 Docket / Notice 编号
     -> 比对已下载文档元数据里的 Docket Number / Notice Number
  B. 修订号线：已下载文档元数据 metadatas["Amendment"]（形如 "23-54|25-100|33-20"）
     -> 直接得出该文档覆盖了哪些 23-xx 修订案
（B 能兜住条款文本里不写案号的情况，例如
 "Final Rule. Amendment 23-61; Published on 6/8/2011."）

输出：FAR23_修订案覆盖报告.csv
"""
import collections
import csv
import glob
import json
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE, "Rulemaking Docs")
CORR = os.path.join(BASE, "FAR23_条款-规则文件对应表.csv")
CLAUSES = os.path.join(BASE, "FAR23_Amdt23-63_条款清单.csv")
OUT = os.path.join(BASE, "FAR23_修订案覆盖报告.csv")


def load_metas():
    out = {}
    for p in glob.glob(os.path.join(DOCS, "_meta", "*.json")):
        try:
            out[os.path.splitext(os.path.basename(p))[0]] = json.load(open(p, encoding="utf-8"))
        except Exception:
            pass
    return out


def norm_num(s):
    s = re.sub(r"(?i)(notice|docket)\s*(no\.?|number)?", "", s or "").strip()
    return re.sub(r"\s+", " ", s).strip(" ,").lower()


def main():
    metas = load_metas()

    # ---- 已下载侧：案号池 + 修订号池 ----
    have_d, have_n = set(), set()
    fr_by_amdt, nprm_by_amdt = collections.defaultdict(set), collections.defaultdict(set)
    for unid, d in metas.items():
        m = d.get("metadatas") or {}
        dt = (m.get("Document Type") or "")
        dn, nn = norm_num(m.get("Docket Number")), norm_num(m.get("Notice Number"))
        if dn: have_d.add(dn)
        if nn: have_n.add(nn)
        amdts = set(re.findall(r"\b(23-\d+[A-Z]?)\b", m.get("Amendment") or ""))
        if not amdts:
            amdts = set(re.findall(r"\b(23-\d+[A-Z]?)\b", str(d.get("docName") or "")))
        if dt.startswith("Final"):
            for a in amdts: fr_by_amdt[a].add(unid)
        elif dt.startswith("Notices"):
            for a in amdts: nprm_by_amdt[a].add(unid)

    def in_pool(kind, num):
        n = norm_num(num)
        if not n:
            return False
        pool = have_n if kind == "NPRM" else have_d
        return any(n == h or n in h or h in n for h in pool)

    # ---- 需求侧：按修订案聚合 ----
    per = collections.defaultdict(lambda: {"fr": set(), "corr": set(), "nprm": set()})
    for r in csv.DictReader(open(CORR, encoding="utf-8-sig")):
        a = (r["Amendment"] or "").strip()
        if not re.fullmatch(r"23-\d+[A-Z]?", a):
            continue
        num = (r["文件编号"] or "").strip()
        if not num:
            continue
        t = (r["文件类型"] or "").upper()
        if "NPRM" in t: per[a]["nprm"].add(num)
        elif "CORRECTION" in t: per[a]["corr"].add(num)
        else: per[a]["fr"].add(num)

    amds = sorted({(c.get("Amendment") or "").strip()
                   for c in csv.DictReader(open(CLAUSES, encoding="utf-8-sig"))
                   if re.fullmatch(r"23-\d+[A-Z]?", (c.get("Amendment") or "").strip())},
                  key=lambda x: (int(re.match(r"23-(\d+)", x).group(1)), x))

    rows = []
    for a in amds:
        frs, cs, nps = sorted(per[a]["fr"]), sorted(per[a]["corr"]), sorted(per[a]["nprm"])

        def status(items, kind, by_amdt):
            if any(in_pool(kind, x) for x in items):
                return "已下载"
            if by_amdt.get(a):
                return "已下载(按修订号匹配)"
            return "DRS无记录" if not items else "未下载"

        rows.append({
            "修订案": a,
            "FinalRule状态": status(frs, "FR", fr_by_amdt),
            "FinalRule案号": "; ".join(frs),
            "更正件状态": status(cs, "FR", fr_by_amdt) if cs else "无",
            "更正件案号": "; ".join(cs),
            "NPRM状态": status(nps, "NPRM", nprm_by_amdt),
            "NPRM案号": "; ".join(nps),
        })

    with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    print("修订案 %d 个 -> %s" % (len(rows), OUT))
    for col in ("FinalRule状态", "更正件状态", "NPRM状态"):
        print("  %-12s %s" % (col, dict(collections.Counter(r[col] for r in rows))))
    left = [r for r in rows if any("未下载" == r[c] for c in
                                   ("FinalRule状态", "更正件状态", "NPRM状态"))]
    print("\n仍未下载: %d" % len(left))
    for r in left:
        print("   %-7s FR:%s %s | 更正:%s %s | NPRM:%s %s" % (
            r["修订案"], r["FinalRule状态"], r["FinalRule案号"],
            r["更正件状态"], r["更正件案号"], r["NPRM状态"], r["NPRM案号"]))


if __name__ == "__main__":
    main()
