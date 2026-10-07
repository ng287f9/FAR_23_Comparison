#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T2 收尾 — 生成 Part 23 更正/校正类文献的权威台账

三个来源合一：
  ① 官方 CFR（govinfo CFR-2016-title14-vol1.xml 抽取的溯源注）
     → FAR23_条款溯源注_官方CFR.csv / FAR23_更正通告清单_官方口径.csv
  ② federalregister.gov API（1994 年后，含 C- 前缀更正件与 CFR Correction）
     → FAR23_1994后Part23规则文件总表.csv
  ③ 本地 PDF 抽取 / govinfo 整期影印（1994 年前无单篇数据的部分）

关键结论（人工核实，写死在 MANUAL 表里）：
  · 1994 年前的更正件 govinfo 只有整期影印 PDF，除已抽的三份外暂不可得
  · 有一批「更正件」其实**不动 Part 23 任何条文**，只修正：
      - 其他部的修订案编号（61 FR 7409 / 61 FR 10269 / 73 FR 65968）
      - 其他部的条文（82 FR 2193 → §91.176；63 FR 53278 → Part 33）
    台账里用「是否影响Part23条文 = 否」标出，避免下游成文时误引

输出：FAR23_更正通告权威台账.csv
"""
import csv
import re
from collections import OrderedDict, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CORR = ROOT / "Rulemaking Docs" / "Corrections"

# ---------------------------------------------------------------- 人工核实表
# FR引证 -> (是否影响Part23条文, 更正对象, 说明)
MANUAL = {
    "30 FR 258": (True, "29 FR 17955 (Amdt 23-0 原始规章)",
                  "§23.365(d) 系数 1.3→1.33；§23.427 公式 'local'→'load'"),
    "58 FR 51970": (True, "58 FR 42136 (Amdt 23-45)",
                    "32 条印刷更正；含 appendix H→appendix I、§23.613/23.865/23.23 等编号错植"),
    "43 FR 52495": (True, "43 FR 50578 (Amdt 23-23)",
                    "§23.785 标题 'Seats and berths.'→'Seats, berths, safety belts, and harnesses.'"
                    "（risingup 溯源注完全漏记此份）"),
    "73 FR 19746": (True, "Amdt 23-45 之后对 §23.573 的再修订",
                    "重写 §23.573(b) 首句——损伤容限评定须含疲劳/腐蚀/意外损伤位置与模式"),
    "73 FR 35063": (True, "§23.1557", "改写 §23.1557(c)(1) 引导句（燃油加油口标记）"),
    "74 FR 32799": (True, "73 FR 12542 (Amdt 23-58)",
                    "§23.1457(d)(1) 与 §23.1459(a)(3) 改为 (i)/(ii) 分项结构"),
    "71 FR 537": (True, "71 FR 首版晾 Final Rule", "§23.773(b) 重写（驾驶舱视野除雾/除霜）"),
    "68 FR 75390": (True, "65 FR 55848 (Amdt 23-54)",
                    "§23.903(a)(2)(i) 引用条款日期更正"),
    "61 FR 252": (True, "Small Airplane Airworthiness Review Program Amdt 23-43",
                  "§23.965(b) 增补 (b)(4)(b)(5) 燃油箱试验振动循环折算"),
    "52 FR 34745": (True, "52 FR 1806 (Amdt 23-34 通勤类最终规章)",
                    "18 项印刷更正；其中第 9 项新增修订指令 25-1："
                    "§23.561(b)(2) 表格第一栏标题 'Normal and utility categories'→"
                    "'Normal, utility, and commuter categories'。"
                    "⚠ 本件 FR 抬头把 Amdt Nos. 误印为 23-24（正文通篇写 23-34），"
                    "CFR 溯源注因此跟着写错 —— DRS 归 Amdt 23-34 是对的"),
    "52 FR 7262": (True, "52 FR 1806 (Amdt 23-34 通勤类最终规章)",
                   "23 项更正；涉条文的是 §23.67(e)(3) 'Vs4'→'Vs1'、"
                   "§23.443(b) 'the'→'The'、§23.1201(a) 'Material'→'No material'，"
                   "其余为前言拼写错误"),
    "71 FR 30577": (True, "CFR 2006 年版 §23.1511",
                    "CFR Correction——删除重复误植的 (a)(2)(i) 与 (a)(2)(ii)"),
    "72 FR 59939": (True, "CFR 2007 年版 §23.561",
                    "CFR Correction——删除重复的第二组 (d)(1)(i)~(iv) 及 (d)(1)(v)"),
    "72 FR 72915": (True, "CFR 2007 年版 §23.561",
                    "CFR Correction（第二次，修正上一条表述）——删除第二段起的五段"),
    "76 FR 81790": (True, "CFR 2011 年版 Appendix C to Part 23",
                    "CFR Correction——Note (4) 更正为 'L is defined in Sec. 23.725(b).'"),
    "67 FR 9552": (True, "Part 23 全书",
                   "清理过期 SFAR：自 Part 23 移除 SFAR No. 41 编者注（Amdt 23-55）"),
    "61 FR 7409": (False, "61 FR 5151 (Docket 27806)",
                   "仅修正 Part 91 修订案编号 91-247→91-248，不动 Part 23 条文"),
    "61 FR 10269": (False, "61 FR 7410 (上一条更正件本身)",
                    "仅把上一条里误写的 '121-248' 改回 '91-248'，不动 Part 23 条文"),
    "73 FR 65968": (False, "73 FR 63339 (Amdt 23-59/35-5)",
                    "仅修正 Amdt 35-5→35-8 与 'S-P'→'CS-P' 标题，不动 Part 23 条文"),
    "82 FR 2193": (False, "81 FR 90126 (Amdt 23-63 EFVS 最终规章)",
                   "C1- 前缀更正件；更正内容在 §91.176(b)(3)(iii)，不动 Part 23 条文"),
    "63 FR 53278": (False, "63 FR 14794 (Docket 28652)",
                    "更正 Part 33 §33.77 吸入条件表与讨论段落，不动 Part 23 条文"),
}

# ------------------------------------------------- 从更正件正文逐条核实的影响面
# 溯源注只记录「该条款所在那一页」的 FR 引证，会漏掉同一份更正件里
# 被前缀 "Amdt. 23-xx" 遮住的其他条款（判定规则把它们当成修订案了）。
# 下表是**逐份读原文**得到的完整清单，优先级高于溯源注推断。
VERIFIED_SECS = {
    "52 FR 34745": "23.3, 23.53, 23.65, 23.67, 23.335, 23.443, 23.561, 23.787, "
                   "23.901, 23.1199, 23.1305, 23.1323, 23.1351, "
                   "Appendix F to Part 23—Test Procedure, "
                   "Appendix G to Part 23—Instructions for Continued Airworthiness",
    "52 FR 7262": "23.67, 23.443, 23.1201",
    "30 FR 258": None,   # 溯源注已全（30 条），不覆盖
}

# CFR 溯源注写的是「条款所在页」，FR API 的 citation 是「文档首页」，
# 同一份文献因此会有两个 FR 引证，这里做别名归并。
ALIAS = {
    "73 FR 35063": "73 FR 35062",   # 文档起于 35062，§23.1557 落在 35063
    "74 FR 32800": "74 FR 32799",   # 同一份 E9-16056 的次页，§23.1459 落在此页
}

KIND_CN = {
    "correction": "更正通告 (FAA Correction)",
    "cfr-correction": "CFR 编纂校正 (OFR CFR Correction)",
    "doc-correction": "更正通告 (FR C- 前缀)",
    "technical-amendment": "技术性修订 (Technical Amendment)",
    "final-rule": "最终规章",
    "special-conditions": "专用条件",
    "(none)": "未分类",
}

# CFR 官方 XML 自身的印刷错误：**FR 卷号 = 年份 - 1935**（1965 年起恒定），
# 据此可抓出溯源注里年份印错的地方。已核实：
#   §23.1105 的 "30 FR 258, Jan. 9, 1996" → 应为 1965-01-09
#   §23.929  的 "33 FR 31822, Nov. 19, 1973" → 应为 1968-11-19
AU_DATE = {"30 FR 258": "⚠ CFR 2016 XML 在 §23.1105 处误印 1996，正确为 1965-01-09"}


def audit_source_note_dates():
    """用「卷号+1935=年份」校验全部 CFR 溯源注，找出官方印刷错误"""
    p = ROOT / "FAR23_条款溯源注_官方CFR.csv"
    if not p.exists():
        return [], []
    bad, checked = [], 0
    pat = re.compile(r"(\d{1,3})\s+FR\s+(\d{1,6}),\s*([A-Za-z]{3,9})\.?\s*(\d{1,2}),\s*(\d{4})")
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        for m in pat.finditer(r.get("官方溯源注") or ""):
            vol, y = int(m.group(1)), int(m.group(5))
            if vol < 30:
                continue
            checked += 1
            if vol + 1935 != y:
                bad.append([r["条款"], m.group(0), str(vol + 1935), m.group(5)])
    return bad, checked


def load_official():
    """官方 CFR 2016 溯源注里得到的更正件清单"""
    p = ROOT / "FAR23_更正通告清单_官方口径.csv"
    out = OrderedDict()
    if not p.exists():
        return out
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        fr = (r.get("更正件FR引证") or "").strip()
        if not fr:
            continue
        out[fr] = {
            "date": (r.get("更正日期") or "").strip(),
            "secs": (r.get("受影响条款") or "").strip(),
            "n": (r.get("影响条款数") or "").strip(),
        }
    return out


def load_api_rows():
    """FR API 1994 后总表里非普通 final-rule/专用条件的条目"""
    p = ROOT / "FAR23_1994后Part23规则文件总表.csv"
    out = OrderedDict()
    if not p.exists():
        return out
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        kind = (r.get("性质") or "").strip()
        if kind in ("final-rule", "special-conditions"):
            continue
        fr = (r.get("FR引证") or "").strip()
        if not fr:
            continue
        out[fr] = {
            "date": (r.get("出版日") or "").strip(),
            "kind": kind,
            "action": (r.get("action") or "").strip(),
            "title": (r.get("标题") or "").strip(),
            "docnum": (r.get("文档号") or "").strip(),
            "secs": (r.get("正文提及条款") or "").strip(),
            "file": (r.get("本地文件") or "").strip(),
            "in_src_note": (r.get("溯源注是否引证") or "").strip(),
        }
    return out


def resolve_local():
    """Corrections/ 目录里现成的 .clean.txt -> FR引证"""
    idx = {}
    for p in CORR.glob("*.clean.txt"):
        m = re.match(r"^(CFRCorrection|Correction|TechAmend|Unclassified)_(\d+)FR(\d+)_",
                     p.name)
        if m:
            idx[f"{m.group(2)} FR {m.group(3)}"] = f"Corrections/{p.name}"
    return idx


def main():
    official = load_official()
    api = load_api_rows()
    local = resolve_local()

    keys = sorted(set(official) | set(api) | set(local),
                  key=lambda k: (official.get(k, {}).get("date")
                                 or api.get(k, {}).get("date") or "9999"))

    rows = []
    for k in keys:
        o, a = official.get(k, {}), api.get(k, {})
        peer = ALIAS.get(k)
        if peer:                       # 别名走主条目的结论与文件
            o = o or official.get(peer, {})
            a = a or api.get(peer, {})
            man = MANUAL.get(peer) or MANUAL.get(k)
        else:
            man = MANUAL.get(k)
        date = o.get("date") or a.get("date") or ""
        if date and k in ALIAS and not o.get("date"):
            date = date
        affects = "是" if man and man[0] else ("否" if man else "")
        vsec = VERIFIED_SECS.get(k)
        secs = vsec or o.get("secs") or a.get("secs", "")
        nsec = o.get("n", "")
        if vsec:
            nsec = str(len([x for x in vsec.split(",") if x.strip()]))
        rows.append(OrderedDict([
            ("FR引证", k),
            ("出版日", date),
            ("性质", KIND_CN.get(a.get("kind", ""), a.get("kind") or "更正通告 (CFR 溯源注)")),
            ("action", a.get("action", "")),
            ("是否影响Part23条文", affects),
            ("影响条款数", nsec),
            ("受影响条款", secs),
            ("更正对象", man[1] if man else ""),
            ("说明", man[2] if man else ""),
            ("CFR溯源注是否列出", a.get("in_src_note", "是" if o else "否")),
            ("标题", a.get("title", "")),
            ("本地文件", local.get(k) or local.get(peer, "") or a.get("file", "")),
            ("本地原文状态", "✅已获取" if (local.get(k) or local.get(peer, "")
                                       or a.get("file"))
             else "⚠未获取（1994 年前仅存 govinfo 整期影印）"),
            ("别名归并至", peer or ""),
            ("出版日校正", AU_DATE.get(k, "")),
        ]))

    write_csv(ROOT / "FAR23_更正通告权威台账.csv", rows)

    got = sum(1 for r in rows if r["本地原文状态"].startswith("✅"))
    hit = sum(1 for r in rows if r["是否影响Part23条文"] == "是")
    print(f"→ FAR23_更正通告权威台账.csv")
    print(f"   共 {len(rows)} 份更正/校正类文献")
    print(f"   本地已获取原文 {got} 份")
    print(f"   实际影响 Part 23 条文 {hit} 份；"
          f"其余 {len(rows)-hit} 份仅修正他部编号/条文，成文时不应误引")
    print("\n   未获取清单（按年份）：")
    for r in rows:
        if not r["本地原文状态"].startswith("✅"):
            print(f"     {r['FR引证']:<14} {r['出版日']}  {r['影响条款数'] or '?':>2} 条"
                  f"  {r['受影响条款'][:70]}")

    write_sec_ledger(official, api, rows)
    write_date_audit()


def write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def write_sec_ledger(official, api, corr_rows):
    """条款级视角：每个条款/附录受哪些更正件影响，正文是否可得"""
    file_of = {r["FR引证"]: r["本地文件"] for r in corr_rows}
    desc_of = {r["FR引证"]: r["说明"] for r in corr_rows}
    hitof = {r["FR引证"]: r["是否影响Part23条文"] for r in corr_rows}

    srcp = ROOT / "FAR23_条款溯源注_官方CFR.csv"
    sec_rows, subpart_rows = [], []
    if srcp.exists():
        for r in csv.DictReader(open(srcp, encoding="utf-8-sig")):
            if int(r.get("更正件数") or 0) <= 0:
                continue
            sec_rows.append([r["条款"], "", ""])
    by_sec = defaultdict(list)
    for s, _, _ in sec_rows:
        pass

    # 直接从台账反查（覆盖面比溯源注更全，含 FR API 新发现）
    sec2corr = defaultdict(set)
    for fr, info in official.items():
        for s in (x.strip() for x in info["secs"].split(",") if x.strip()):
            sec2corr[s].add(fr)
    for fr, info in api.items():
        for s in (x.strip() for x in info.get("secs", "").split(",") if x.strip()):
            sec2corr[s].add(fr)
    for fr, v in VERIFIED_SECS.items():
        if v:
            for s in (x.strip() for x in v.split(",") if x.strip()):
                sec2corr[s].add(fr)
    for fr, secs in [("43 FR 52495", "23.785"), ("61 FR 252", "23.965"),
                     ("67 FR 9552", "Part 23 全书（Part 23 序言区）"),
                     ("71 FR 30577", "23.1511"), ("72 FR 59939", "23.561"),
                     ("72 FR 72915", "23.561"), ("76 FR 81790", "Appendix C to Part 23"),
                     ("73 FR 35063", "23.1557"), ("74 FR 32800", "23.1459")]:
        sec2corr[secs].add(fr)

    def key(s):
        m = re.match(r"^23\.(\d+)", s)
        if m:
            return (0, int(m.group(1)))
        return (1, 0)

    out = []

    def sort_key(s):
        m = re.match(r"^23\.(\d+)", s)
        n = int(m.group(1)) if m else 10**6
        return (n if m else 10**6, s)

    for s in sorted(sec2corr, key=sort_key):
        frs = sorted(sec2corr[s])
        have = sum(1 for f in frs if file_of.get(f))
        out.append(OrderedDict([
            ("条款/附录", s),
            ("受影响更正件数", len(frs)),
            ("更正件FR引证", "; ".join(frs)),
            ("原文已获取", f"{have}/{len(frs)}"),
            ("更正内容要点", " | ".join(desc_of.get(f, "") for f in frs
                                    if desc_of.get(f) and hitof.get(f) == "是")[:400]),
        ]))
    write_csv(ROOT / "FAR23_受影响条款清单.csv", out)
    print(f"→ FAR23_受影响条款清单.csv  ({len(out)} 个条款/附录受影响)")


def write_date_audit():
    bad, checked = audit_source_note_dates()
    rows = [OrderedDict([("条款", b[0]), ("原文引证", b[1]),
                         ("应为年份", b[2]), ("官方误印年份", b[3])]) for b in bad]
    if rows:
        write_csv(ROOT / "FAR23_溯源注日期勘误.csv", rows)
        print(f"→ FAR23_溯源注日期勘误.csv  "
              f"（校验 {checked} 条 FR 引证，发现 {len(bad)} 处 CFR 官方印刷错误）")
        for b in bad:
            print(f"     §{b[0]}: {b[1]}  应为 {b[2]}，官方印 {b[3]}")
    else:
        print(f"   日期校验通过（{checked} 条 FR 引证）")


if __name__ == "__main__":
    main()
