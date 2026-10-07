#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 govinfo 的 CFR 年度版 XML 提取 **14 CFR Part 23 全部条款的官方溯源注**，
判定哪些条款的来源注里挂着「裸 FR 引证」＝ 更正通告（Correction），
并与 risingup.com 缓存页面逐条比对，标出两者的分歧。

输出：
    FAR23_条款溯源注_官方CFR.csv     每条款：官方溯源注 + 更正件判定
    FAR23_更正通告清单_官方口径.csv    按更正件汇总（FR引证/日期/影响条款）
    FAR23_risingup与官方CFR差异.csv   risingup 与官方不一致的条款

用法：
    python scan_corrections_official.py [--year 2016]
"""
import csv
import html
import os
import re
import subprocess
import sys
import urllib.request
import gzip
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TMP = Path("/tmp")
RISING = Path("/tmp/risingup_pages")

CFR_URL = ("https://www.govinfo.gov/content/pkg/"
           "CFR-{y}-title14-vol1/xml/CFR-{y}-title14-vol1.xml")

AMDT_RE = re.compile(r"Amdt\.?\s*(?:No\.?\s*)?\s*(\d+)\s*[‐-―\-−]\s*(\d+[A-Z]?)")
DOCNO_RE = re.compile(r"^\[?\s*(?:Doc(?:ket)?\.?\s*No\.?|Docket)\s*", re.I)
FR_RE = re.compile(r"(\d{1,3})\s+FR\s+(\d{1,6})")
DATE_RE = re.compile(
    r"(Jan\.|Feb\.|Mar\.|Apr\.|May|June|July|Aug\.|Sept?\.|Oct\.|Nov\.|Dec\.)"
    r"\s+(\d{1,2}),\s+(\d{4})", re.I)

MONTHS = {"jan.": 1, "feb.": 2, "mar.": 3, "apr.": 4, "may": 5, "june": 6,
          "july": 7, "aug.": 8, "sep.": 9, "sept.": 10 - 1, "oct.": 10,
          "nov.": 11, "dec.": 12}


def fetch(year: int) -> Path:
    p = TMP / f"CFR-{year}-title14-vol1.xml"
    if p.exists() and p.stat().st_size > 1_000_000:
        return p
    url = CFR_URL.format(y=year)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0",
                                               "Accept-Encoding": "gzip"})
    with urllib.request.urlopen(req, timeout=120) as r, open(p, "wb") as f:
        raw = r.read()
        if r.headers.get("Content-Encoding") == "gzip":
            raw = gzip.decompress(raw)
        f.write(raw)
    return p


def find_part23(xml_path: Path):
    r = ET.parse(xml_path).getroot()
    for e in r.iter("PART"):
        for ch in list(e)[:5]:
            if "".join(ch.itertext()).strip().startswith("PART 23"):
                return e
    return None


def iter_sections(part):
    for sec in part.iter("SECTION"):
        no = sec.find("SECTNO")
        if no is None:
            continue
        num = "".join(no.itertext()).replace("§", "").strip()
        subj_e = sec.find("SUBJECT")
        subj = "".join(subj_e.itertext()).strip() if subj_e is not None else ""
        cita = sec.find("CITA")
        note = re.sub(r"\s+", " ", "".join(cita.itertext())).strip() if cita is not None else ""
        yield num, subj, note
    # 附录（Appendix A–J）：附录整体也带一条溯源注
    for app in part.iter("APPENDIX"):
        hd = app.find("HD") if app.find("HD") is not None else app.find("HEAD")
        name = "".join(hd.itertext()).strip() if hd is not None else ""
        name = name.replace("Appendix to Part 23", "Appendix").replace(
            "Appendix to Part", "Appendix").strip()
        cita = app.find("CITA")
        note = re.sub(r"\s+", " ", "".join(cita.itertext())).strip() if cita is not None else ""
        if not cita:
            for ch in app:
                if ch.tag == "CITA":
                    note = re.sub(r"\s+", " ", "".join(ch.itertext())).strip()
                    break
        if note:
            yield name, "（附录）", note


def split_clauses(note: str):
    """溯源注切句：既按 ';' 切，也按 'as amended by' 切。"""
    parts = []
    for chunk in re.split(r"as amended by", note, flags=re.I):
        for c in chunk.split(";"):
            c = c.strip().strip("[]").strip()
            if c:
                parts.append(c)
    return parts


def classify(note: str):
    """返回 (FR引证总数, 修订案列表, 更正件列表)

    更正件判定 = 「裸 FR 引证」且**前面已经出现过一次 FR 引证**。
    加这个前置条件是为了排除 1996 年后 OCR 化的新式溯源注
    `[Doc. No. FAA-2010-0224; 76 FR 33135, June 8, 2011]`——那里 FR 是主引证，
    不是更正件。
    """
    if not note:
        return 0, [], []
    note_clean = note.strip().lstrip("[").rstrip("]")
    if DOCNO_RE.match(note_clean):
        note_clean = DOCNO_RE.sub("", note_clean, count=1)
    clauses = split_clauses(note_clean)

    amends, corrs, seen_fr = [], [], False
    for c in clauses:
        c = c.strip()
        amends += AMDT_RE.findall(c)
        has_fr = FR_RE.search(c) is not None
        m = re.match(r"^\s*(\d{1,3})\s+FR\s+(\d{1,6})", c)
        if m and seen_fr:  # 之前已经出现过一次 FR 引证 → 这是更正件
            vol, page = m.groups()
            d = DATE_RE.search(c)
            date = ""
            if d:
                mo = MONTHS.get(d.group(1).lower().replace("sept.", "sep."), 0)
                date = f"{int(d.group(3)):04d}-{mo:02d}-{int(d.group(2)):02d}"
            corrs.append((f"{vol} FR {page}", date))
        seen_fr = seen_fr or has_fr
    fr_total = len(FR_RE.findall(note))
    seen, uniq = set(), []
    for c in corrs:
        if c[0] not in seen:
            seen.add(c[0])
            uniq.append(c)
    return fr_total, amends, uniq


def load_risingup():
    """从本地缓存读取 risingup 溯源注（无缓存返回 None）"""
    if not RISING.is_dir():
        return None
    out = {}
    for f in RISING.glob("*.shtml"):
        t = f.read_text(encoding="utf-8", errors="replace")
        t = re.sub(r"<script.*?</script>", "", t, flags=re.S | re.I)
        t = re.sub(r"<[^>]+>", " ", t)
        t = html.unescape(t)
        t = re.sub(r"\s+", " ", t)
        m = re.search(r"(\d{2}\s+FR\s+\d{1,6})", t)
        num = f.stem.replace("_fars_info_part23-", "").replace("-FAR", "")
        # 取第一条形如 [Doc. No. ... ] 的溯源注
        n = re.search(r"\[[Dd]oc(?:ket)?\.?\s*[Nn]o\.?.{0,700}?\]", t)
        out[num] = (n.group(0) if n else "")
    return out


def main():
    year = 2016
    if "--year" in sys.argv:
        year = int(sys.argv[sys.argv.index("--year") + 1])
    xml = fetch(year)
    print(f"CFR {year} XML: {xml}  ({xml.stat().st_size:,} bytes)")
    part = find_part23(xml)
    assert part is not None, "找不到 PART 23"

    rows = []
    for num, subj, note in iter_sections(part):
        fr_total, amends, corrs = classify(note)
        rows.append({
            "条款": num,
            "标题": subj,
            "官方溯源注": note,
            "FR引证数": fr_total,
            "修订案": "; ".join(f"{a}-{b}" for a, b in amends),
            "更正件数": len(corrs),
            "更正件FR引证": "; ".join(f"{c[0]} ({c[1]})" for c in corrs),
        })
    print(f"Part 23 条款数 {len(rows)}")
    hit = [r for r in rows if r["更正件数"]]
    print(f"带更正件引证的条款 {len(hit)}")

    out1 = ROOT / "FAR23_条款溯源注_官方CFR.csv"
    with open(out1, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"→ {out1.name}")

    # 按更正件汇总（并做三件后处理）
    agg = {}
    for r in rows:
        for c in (x.strip() for x in r["更正件FR引证"].split(";") if x.strip()):
            agg.setdefault(c, []).append(r["条款"])
    from collections import OrderedDict

    # 1) 官方 CFR 把 §23.1105 的 30 FR 258 日期印成 1996（应为 1965，30 FR 就是 1965 卷）
    #    → 合并到同一份更正件
    for k in list(agg):
        if k.startswith("30 FR 258") and "1996" in k:
            tgt = "30 FR 258 (1965-01-09)"
            agg.setdefault(tgt, []).extend(agg.pop(k))
    # 2) 本地正文可得性
    LOCAL = {
        "30 FR 258": "✅ Corrections/Correction_30FR258_1965-01-09_Amdt23-0.clean.txt"
                     "（本轮从 govinfo FR-1965-01-09 整期 PDF 抽 OCR 文字层）",
        "58 FR 51970": "✅ Corrections/Correction_58FR51970_1993-10-05_Amdt23-45.clean.txt"
                       "（32 条更正项，Local PDF 抽出）",
        "71 FR 537": "✅ Corrections/Correction_71FR537_2006-01-05_06-85.clean.txt",
        "73 FR 19746": "✅ Corrections/Correction_73FR19746_2008-04-11_E8-7649.clean.txt",
        "73 FR 35063": "✅ Corrections/Correction_73FR35063_2008-06-20_E8-13900.clean.txt",
        "74 FR 32799": "✅ Corrections/Correction_74FR32799_2009-07-09_E9-16056.clean.txt",
        "74 FR 32800": "✅ 与上一条同一份文档 E9-16056（74 FR 32799-32802）",
    }
    # 3) govinfo 整期 PDF 体积（用于估算补抓成本）
    ISSUE_MB = {"32 FR 13505": ("1967-09-27", 15), "32 FR 13714": ("1967-09-30", 11),
                "34 FR 14727": ("1969-09-24", 6), "34 FR 17509": ("1969-10-30", 15),
                "35 FR 1102": ("1970-01-28", 19), "38 FR 32784": ("1973-11-28", 24),
                "52 FR 7262": ("1987-03-09", 44), "52 FR 34745": ("1987-09-14", 38),
                "53 FR 34194": ("1988-09-02", 70), "55 FR 46888": ("1990-11-07", 168),
                "56 FR 5455": ("1991-02-11", 50), "58 FR 18975": ("1993-04-09", 504),
                "58 FR 27060": ("1993-05-06", 316)}

    agg = OrderedDict(sorted(agg.items(), key=lambda kv: -len(kv[1])))
    out2 = ROOT / "FAR23_更正通告清单_官方口径.csv"
    with open(out2, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["更正件FR引证", "更正日期", "影响条款数", "受影响条款", "本地正文", "补抓成本"])
        for k, v in agg.items():
            cite = k.split(" (")[0]
            date = k.split(" (")[1].rstrip(")") if " (" in k else ""
            if cite in LOCAL:
                loc, cost = LOCAL[cite], ""
            elif cite in ISSUE_MB:
                d, mb = ISSUE_MB[cite]
                loc = "⚠ 未获取"
                cost = f"govinfo FR-{d} 整期 PDF ~{mb}MB（有 OCR 文字层，可行）"
            else:
                loc, cost = "⚠ 未获取", ""
            w.writerow([cite, date, len(v), ", ".join(sorted(v)), loc, cost])
    print(f"→ {out2.name}（{len(agg)} 份独立更正件，覆盖 {sum(len(v) for v in agg.values())} 条款次）")

    # 与 risingup 比对
    ru = load_risingup()
    if ru:
        diffs = []
        for r in rows:
            num = r["条款"].replace("§", "").strip()
            # risingup 文件名形如 part23-525-FAR.shtml / part23-A-APPX.shtml
            tail = num.split(".")[1] if "." in num else num
            cand = ru.get(tail)
            if cand is None:
                continue
            a = re.sub(r"\s+", " ", r["官方溯源注"]).strip()
            b = re.sub(r"\s+", " ", cand).strip()
            # 只看 risingup 里是否多出官方没有的 FR 页码
            extra = sorted(set(FR_RE.findall(b)) - set(FR_RE.findall(a)))
            if extra:
                diffs.append({
                    "条款": num,
                    "risingup多出的FR引证": ", ".join(f"{v} FR {p}" for v, p in extra),
                    "官方CFR溯源注": a,
                    "risingup溯源注": b,
                })
        out3 = ROOT / "FAR23_risingup与官方CFR差异.csv"
        with open(out3, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(diffs[0].keys()) if diffs
                               else ["条款", "risingup多出的FR引证", "官方CFR溯源注", "risingup溯源注"])
            w.writeheader()
            w.writerows(diffs)
        print(f"→ {out3.name}（{len(diffs)} 条分歧）")


if __name__ == "__main__":
    main()
