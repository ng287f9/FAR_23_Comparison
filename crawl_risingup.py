#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
抓取 risingup.com 上 Part 23 每个条款页面，解析 CFR 条款溯源注（source credit），
识别"同一 Doc. No. 下出现第二个 FR 引证"＝更正通告（correction）的情形。

输出：/Users/glennchou/FAR 23/FAR23_条款溯源注.csv
"""
import csv
import html
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor

import requests

BASE = "https://www.risingup.com"
INDEX = BASE + "/fars/info/23-index.shtml"
OUT = "/Users/glennchou/FAR 23/FAR23_条款溯源注.csv"
CACHE = "/tmp/risingup_pages"

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

os.makedirs(CACHE, exist_ok=True)


def get(url, tries=3):
    for i in range(tries):
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
            if r.status_code == 200:
                return r.text
        except requests.RequestException:
            pass
        time.sleep(1 + i)
    return None


def page_to_text(h):
    h = re.sub(r"(?is)<(script|style|nav|footer|header)\b.*?</\1>", " ", h)
    h = re.sub(r"(?i)<br\s*/?>", "\n", h)
    h = re.sub(r"(?i)</(p|div|li|h[1-6]|tr)>", "\n", h)
    h = re.sub(r"(?s)<[^>]+>", "", h)
    h = html.unescape(h)
    lines = [re.sub(r"[ \t\xa0]+", " ", l).strip() for l in h.split("\n")]
    return [l for l in lines if l]


FR = re.compile(r"(\d{1,3})\s+FR\s+(\d{1,6})\s*,\s*([A-Z][a-z]+\.?)\s+(\d{1,2}),\s*(\d{4})")
# 注意：站点用 en dash（– U+2013）而不是 ASCII 连字符，必须一并接受。
# 且实际写法有 "Amdt. 23–7" / "Amdt. No. 23–59" / "Amdt 23–34" 三种，必须全认。
AMDT = re.compile(r"Amdt\.?\s*(?:No\.?\s*)?\s*(\d+)\s*[‐-―\-−]\s*(\d+[A-Z]?)", re.I)
SFAR = re.compile(r"SFAR\s*(\d+)\s*[‐-―\-−]\s*(\d+[A-Z]?)", re.I)
DOCNO = re.compile(r"Doc\.\s*No\.\s*:?\s*([^,;]+)", re.I)
DOCKET = re.compile(r"Docket\s*(?:No\.?)?\s*(FAA-[\d-]+|\d+)", re.I)
MONTH = {m: i + 1 for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}


def parse_credit(lines):
    """从页面文本里找到 [...] 溯源注，并解析出 FR 引证序列。"""
    credit = None
    for l in lines:
        if l.startswith("[") and ("FR" in l or "Amdt" in l or "Doc. No" in l):
            # 取最长的一条（有些页面会拆成多行）
            credit = l
            break
    if not credit:
        # 兜底：把相邻的 [ ... ] 拼起来
        joined = " ".join(lines)
        m = re.search(r"\[[^\[\]]{20,1200}?\]", joined)
        if m and "FR" in m.group(0):
            credit = m.group(0)
    if not credit:
        return None, []
    items = []
    # 关键：不仅按分号切，还要按 "as amended by" 切。
    # 形如 "…; 58 FR 51970, Oct. 5, 1993, as amended by Amdt. 23–48, 61 FR 5147, Feb. 9, 1996"
    # 这一整段里既有更正件（裸 FR）又有修订案，必须拆开，否则裸 FR 会被 Amdt 标记"带偏"而漏判。
    for chunk in re.split(r";", credit):
        for seg in re.split(r"as\s+amended\s+by", chunk, flags=re.I):
            seg = seg.strip(" []")
            if not seg:
                continue
            mfr = FR.search(seg)
            if not mfr:
                continue
            ma = AMDT.search(seg) or SFAR.search(seg)
            md = DOCNO.search(seg) or DOCKET.search(seg)
            mon = MONTH.get(mfr.group(3)[:3], 0)
            items.append({
                "seg": seg.strip(),
                "vol": int(mfr.group(1)),
                "page": int(mfr.group(2)),
                "date": f"{mfr.group(5)}-{mon:02d}-{int(mfr.group(4)):02d}",
                "amdt": (f"{ma.group(1)}-{ma.group(2)}" if ma else ""),
                "docno": (md.group(1).strip() if md else ""),
            })
    return credit, items


def title_of(lines):
    for l in lines:
        m = re.match(r"Sec\.\s*([\w.]+)\s*[-—–]\s*(.+?)\.?$", l)
        if m:
            return m.group(1), m.group(2)
    return "", ""


def main():
    idx = get(INDEX)
    links = sorted(set(re.findall(r'href="([^"]*part23-[^"]*\.shtml)"', idx or "")))
    print("待抓取页面:", len(links))

    def work(rel):
        url = BASE + rel if rel.startswith("/") else BASE + "/fars/info/" + rel
        fn = os.path.join(CACHE, rel.replace("/", "_"))
        if os.path.exists(fn):
            h = open(fn, encoding="utf-8", errors="replace").read()
        else:
            h = get(url)
            if not h:
                return None
            open(fn, "w", encoding="utf-8").write(h)
            time.sleep(0.15)
        lines = page_to_text(h)
        sec, title = title_of(lines)
        credit, items = parse_credit(lines)
        return {"rel": rel, "url": url, "sec": sec, "title": title,
                "credit": credit or "", "items": items}

    rows = []
    with ThreadPoolExecutor(max_workers=6) as ex:
        for r in ex.map(work, links):
            if r:
                rows.append(r)
    rows.sort(key=lambda r: r["rel"])
    print("抓取成功:", len(rows))

    import part23_source as S
    doc_index = S.build_doc_index()

    # ---- 已知"修订案"文献的 FR 页码表（用于排除"溯源注漏写 Amdt 标记"的假阳性）----
    amd_fr = []
    for d in doc_index.values():
        if d.kind == "FinalRule" and d.fr:
            amd_fr.append((d.fr[0], d.fr[1], d))
    # DRS 条款清单里带 Amdt 标记的页码也算修订案
    for r in rows:
        for it in r["items"]:
            if it["amdt"]:
                amd_fr.append((it["vol"], it["page"], None))

    def near_amendment(vol, page, tol=40):
        v = [a for a in amd_fr if a[0] == vol and a[1] and abs(a[1] - page) <= tol]
        return v[0][2] if v else None

    def find_local(vol, page, tol=25):
        best = None
        for d in doc_index.values():
            if d.fr and d.fr[0] == vol:
                if d.fr[1] and abs(d.fr[1] - page) <= tol:
                    return d, "精确"
                if best is None:
                    best = d
        if best is not None:
            return best, "同卷近似"
        return None, ""

    def strict_find(vol, page):
        """只在文件名里真的出现 'vol FR page'（或裸的 >=4 位页码）才算命中。
        模糊的『同卷近似』会大量误报『已有』，不可用。"""
        import os
        pat = re.compile(rf"(?<![0-9]){vol}\s*FR\s*{page}(?![0-9])")
        pat2 = re.compile(rf"(?<![0-9]){page}(?![0-9])") if page >= 1000 else None
        for d in doc_index.values():
            base = os.path.basename(d.path)
            if pat.search(base):
                return d, "文件名含FR引证"
            if pat2 and pat2.search(base):
                return d, "文件名含页码"
        return None, ""

    # 判定"更正"：没有 Amdt 标记的 FR 引证，且不在任何已知修订案文献页码附近
    multi = []
    for r in rows:
        its = r["items"]
        correction = []
        if its:
            # 第一条（Doc. No. 起始那条）不参与判定
            for i in its[1:]:
                if i["amdt"]:
                    continue
                host = near_amendment(i["vol"], i["page"])
                if host is None:
                    i["local"], i["how"] = find_local(i["vol"], i["page"])
                    correction.append(i)
        r["corrections"] = correction
        if correction:
            multi.append(r)

    with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["页面", "条款", "标题", "溯源注原文", "FR引证数",
                    "更正件数", "更正月引证", "本地文件", "URL"])
        for r in rows:
            cs = " ｜ ".join(f"{c['vol']} FR {c['page']} ({c['date']})"
                             for c in r["corrections"])
            locs = []
            for c in r["corrections"]:
                d, how = strict_find(c["vol"], c["page"])
                c["local"], c["how"] = d, how
                locs.append(f"{d.name}[{how}]" if d else "⚠ 未下载")
            w.writerow([r["rel"], r["sec"], r["title"], r["credit"],
                        len(r["items"]), len(r["corrections"]), cs,
                        " ｜ ".join(locs), r["url"]])

    # ---- 汇总：把同一份更正件按 (卷,页) 合并，输出影响条款清单 ----
    merged = {}
    for r in rows:
        for c in r["corrections"]:
            key = (c["vol"], c["page"], c["date"])
            m = merged.setdefault(key, {"secs": [], "local": None, "how": ""})
            m["secs"].append(r["sec"] or r["rel"])
            if m["local"] is None and c.get("local"):
                m["local"], m["how"] = c["local"], c["how"]

    OUT2 = "/Users/glennchou/FAR 23/FAR23_更正通告清单.csv"
    with open(OUT2, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["更正件FR引证", "更正日期", "年代", "影响条款数", "受影响条款", "本地原文"])
        for (vol, page, date) in sorted(merged, key=lambda k: (k[2], k[0], k[1])):
            m = merged[(vol, page, date)]
            era = "1994后(可单篇获取)" if date >= "1994-01-01" else "1994前(仅整期PDF)"
            loc = f"{m['local'].name}[{m['how']}]" if m["local"] else "⚠ 未下载"
            w.writerow([f"{vol} FR {page} ({date})", date, era,
                        len(m["secs"]), ", ".join(sorted(set(m["secs"]))), loc])
    print(f"✅ 已写出: {OUT2}  （{len(merged)} 份独立更正件）")

    print(f"\n✅ 已写出: {OUT}")
    print(f"含更正通告的条款数: {len(multi)} / {len(rows)}")
    for r in multi:
        print(f"  {r['sec']:10s} {r['title'][:40]:42s} | "
              + " ; ".join(f"{c['vol']}FR{c['page']} {c['date']}" for c in r["corrections"]))


if __name__ == "__main__":
    main()
