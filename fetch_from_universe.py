#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 FAR23_Part23规则文件全集.csv 补抓尚未下载的 NPRM / Final Rule。"""
import csv
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fetch_rulemaking as F  # noqa: E402

UNIVERSE = "/Users/glennchou/FAR 23/FAR23_Part23规则文件全集.csv"


def norm_date(s):
    """MM/DD/YYYY -> YYYY-MM-DD"""
    s = (s or "").strip()
    if "/" in s:
        p = s.split("/")
        if len(p) == 3 and len(p[2]) == 4:
            return f"{p[2]}-{p[0].zfill(2)}-{p[1].zfill(2)}"
    return s


def main():
    workers = int(sys.argv[1]) if len(sys.argv) > 1 else 4
    rows = list(csv.DictReader(open(UNIVERSE, encoding="utf-8-sig")))
    have = set(os.path.splitext(f)[0] for f in os.listdir(F.META))

    todo = []
    for r in rows:
        unid = (r.get("unid") or "").upper()
        if not unid or unid in have:
            continue
        num = ("Docket " + r["docket"]) if r["docket"] else ("Notice " + r["notice"])
        todo.append({
            "类型": "Final Rule" if r["kind"] == "FinalRule" else "NPRM",
            "编号": num.strip(),
            "签发日期": norm_date(r.get("issueDate")),
            "unid": unid,
        })

    print(f"全集 {len(rows)} 条，已下载 {len(rows)-len(todo)}，待补 {len(todo)}", flush=True)
    results = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for i, res in enumerate(ex.map(F.fetch_one, todo), 1):
            results.append(res)
            print("[%3d/%d] %s %-30s %s" % (
                i, len(todo), "OK " if res["ok"] else "FAIL", res["编号"][:30],
                ", ".join(res.get("files", []))[:90] or res.get("why", "")), flush=True)

    ok = sum(1 for r in results if r["ok"])
    print(f"\n补抓完成: {ok}/{len(todo)}")
    for r in results:
        if not r["ok"]:
            print("  仍在失败:", r["编号"], r["why"])


if __name__ == "__main__":
    main()
