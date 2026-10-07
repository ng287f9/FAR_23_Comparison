#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 Rulemaking Docs/ 里的文件名改成带**真实《联邦公报》引证**的形式。

为什么要改：原文件名里的日期取自 DRS 的 `Issue Date`（签发日），
与《联邦公报》实际出版日经常差几天到几周，用户按 FR 去查会找不到。
例如 23-61：DRS Issue Date 2011-05-20，实际 76 FR 33129 出版于 2011-06-08。

命名规则：
    {FinalRule|NPRM}_{Docket xxx|Notice xx-yy}[_Amdt..]_{卷}FR{页}_{FR出版日}[_000N].{ext}
    例：FinalRule_Docket FAA-2010-0224_Amdt23-61_76FR33129_2011-06-08.html
卷/页/日期从文档自带的 DRS 头（.txt 或 HTML 里的 CITATION / PAGE NUMBER 字段）提取，
取不到则退回元数据里的 Citation（只有卷和日期，页码留空）。

用法：python rename_with_fr_citation.py [--apply]
不带 --apply 只干跑，打印改名计划。
"""
import glob
import html as _html
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE, "Rulemaking Docs")
META = os.path.join(DOCS, "_meta")

MON = {m.lower(): i + 1 for i, m in enumerate(
    "January February March April May June July August September October November December".split())}


def strip_html(s):
    s = re.sub(r"(?is)<!--.*?-->", " ", s)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?s)<[^>]+>", " ", s)
    return _html.unescape(s)


def parse_fr(text):
    """从一段文本里抽 (卷号, 页码, FR出版日)"""
    vol = page = date = ""
    m = re.search(r"Volume\s+(\d+)", text, re.I)
    if m:
        vol = m.group(1)
    m = re.search(r"\[?Pages?\s+([\d\-]+)\]?", text, re.I)
    if m:
        page = m.group(1).split("-")[0]
    m = re.search(r"([A-Z][a-z]+)\s+(\d{1,2})\s*[,.]\s*(\d{4})", text)
    if m and m.group(1).lower() in MON:
        date = "%s-%02d-%02d" % (m.group(3), MON[m.group(1).lower()], int(m.group(2)))
    return vol, page, date


def base_and_suffix(fn):
    stem, ext = os.path.splitext(fn)
    m = re.match(r"^(.*)_(\d{4})$", stem)
    if m:
        return m.group(1), "_" + m.group(2), ext
    return stem, "", ext


def extract_from_files(files):
    """files: 该 base 下的全部文件名"""
    # 优先 .txt（已经是纯文本，字段干净）
    for pref in (".txt", ".html", ".htm"):
        for fn in sorted(files):
            if not fn.endswith(pref):
                continue
            p = os.path.join(DOCS, fn)
            try:
                s = open(p, encoding="utf-8", errors="replace").read(400000)
            except Exception:
                continue
            t = s if pref == ".txt" else strip_html(s)
            i = t.upper().find("CITATION")
            if i < 0:
                continue
            seg = t[i:i + 300]
            j = t.upper().find("PAGE NUMBER")
            if j >= 0:
                seg += " " + t[j:j + 120]
            vol, page, date = parse_fr(seg)
            if date:
                return vol, page, date
    return "", "", ""


def meta_citation(base):
    """退回元数据 Citation"""
    for p in glob.glob(os.path.join(META, "*.json")):
        try:
            d = json.load(open(p, encoding="utf-8"))
        except Exception:
            continue
        m = d.get("metadatas") or {}
        if not m:
            continue
        dn = (m.get("Docket Number") or "").strip()
        nn = (m.get("Notice Number") or "").strip()
        hit = (dn and ("Docket " + dn) in base) or (nn and ("Notice " + nn) in base)
        if hit:
            vol, page, date = parse_fr(m.get("Citation") or "")
            if date:
                return vol, page, date
    return "", "", ""


def main():
    apply = "--apply" in sys.argv
    files = [f for f in os.listdir(DOCS)
             if os.path.isfile(os.path.join(DOCS, f)) and not f.startswith(("_", "."))]

    groups = {}
    for fn in files:
        b, suf, ext = base_and_suffix(fn)
        groups.setdefault(b, []).append((fn, suf, ext))

    plan = []
    for b, items in sorted(groups.items()):
        vol, page, date = extract_from_files([x[0] for x in items])
        if not date:
            vol, page, date = meta_citation(b)
        if not date:
            plan.append((b, b, "无法确定 FR 日期，保持原名"))
            continue
        # 把 base 里原有的日期（若有）换成 FR 出版日，并插入卷/页
        b2 = re.sub(r"_\d{4}-\d{2}-\d{2}$", "", b)
        b2 = re.sub(r"_\d{1,3}FR\d*$", "", b2)
        cite = ("%sFR%s" % (vol, page)) if vol else ""
        newb = "_".join(x for x in (b2, cite, date) if x)
        plan.append((b, newb, ""))

    print("共 %d 个文档组\n" % len(groups))
    renamed = skipped = 0
    for old, new, note in plan:
        if note:
            print("  [跳过] %-70s %s" % (old[:70], note))
            skipped += 1
            continue
        if old == new:
            continue
        print("  %s\n      -> %s" % (old, new))
        if apply:
            for fn, suf, ext in groups[old]:
                dst = new + suf + ext
                src_p = os.path.join(DOCS, fn)
                dst_p = os.path.join(DOCS, dst)
                if os.path.exists(dst_p) and dst_p != src_p:
                    print("      !! 目标已存在，跳过 %s" % dst)
                    continue
                os.rename(src_p, dst_p)
        renamed += 1
    print("\n%s: 计划改名 %d 组，跳过 %d 组" % ("已执行" if apply else "干跑", renamed, skipped))


if __name__ == "__main__":
    main()
