#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 §23.561 精写模板批量生成《14 CFR Part 23 条款修订历史与背景分析》分部卷。

结构（每个条款严格五节 + 一附录，与 FAR23_Sec23.561_修订历史与背景分析.docx 对齐）：

    § 23.xxx　中文标题（English Title）                       ← 大纲 1 级
        1　§23.xxx 修订历史汇总                               ← 大纲 2 级
            1.1　条款基本信息 / 1.2　修订沿革总表 / 1.3　修订脉络一览
        2　各阶段修订背景与原因深度剖析
            2.1…2.n　每个版本一节（初始颁行 / Amdt xx）
            2.n+1　更正通告与 CFR 编纂校正（有则输出）
        3　条款原文演变对比
            3.1…3.n　逐版条文（段落 | 英文原文 | 状态着色）
            3.n+1　段落结构演变矩阵
        4　总结
        5　参考文献
        附录　DRS 条款版本记录 ID

语言：除条款原文（英文条文、英文引证）与文件名外，全部中文；专业术语中英对照。

用法：
    python build_clause_doc.py C                 # C 分部全卷
    python build_clause_doc.py C --limit 3       # 只做前 3 条（调试）
    python build_clause_doc.py C --no-translate  # 跳过翻译（讨论段落留英文）
"""
import argparse
import csv
import os
import re
import sys
from collections import OrderedDict, Counter, defaultdict

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

import docx_layout as L
import part23_source as S

ROOT = os.path.dirname(os.path.abspath(__file__))

SUBPART_KEY = {
    "A": "Subpart A - General",
    "B": "Subpart B - Flight",
    "C": "Subpart C - Structure",
    "D": "Subpart D - Design and Construction",
    "E": "Subpart E - Powerplant",
    "F": "Subpart F - Equipment",
    "G": "Subpart G - Operating Limitations and Information",
    "AppA": "Appendix A--Simplified Design Load Criteria",
    "AppB": "Appendix B--Control Surface Loadings",
    "AppC": "Appendix C--Basic Landing Conditions",
    "AppD": "Appendix D--Wheel Spin-Up Loads",
    "AppE": "Appendix E--[Removed and Reserved.]",
    "AppF": "Appendix F--Test Procedure",
    "AppG": "Appendix G--Instructions for Continued Airworthiness",
    "AppH": "Appendix H--Installation of an Automatic Power Reserve (APR) System",
    "AppI": "Appendix I--Seaplane Loads",
    "AppJ": "Appendix J--HIRF Envirnoments and Equipment HIRF Test Levels",
}
SUBPART_CN = {
    "A": "A 分部　总则", "B": "B 分部　飞行", "C": "C 分部　结构",
    "D": "D 分部　设计与构造", "E": "E 分部　动力装置", "F": "F 分部　设备",
    "G": "G 分部　使用限制与资料",
    "AppA": "附录 A　简化设计载荷准则", "AppB": "附录 B　操纵面载荷",
    "AppC": "附录 C　基本着陆情况", "AppD": "附录 D　机轮起旋载荷",
    "AppE": "附录 E　[已删并预留]", "AppF": "附录 F　试验程序",
    "AppG": "附录 G　持续适航文件", "AppH": "附录 H　自动功率储备（APR）系统装机",
    "AppI": "附录 I　水上飞机载荷", "AppJ": "附录 J　HIRF 环境与设备 HIRF 试验等级",
}

AMDT_CN = {"Initial": "初始版（Initial）"}


def amdt_cn(a):
    a = (a or "").strip()
    if a.lower() in ("initial", "unnamed", ""):
        return "初始版"
    return f"Amdt {a}"


# ══════════════════════════════════════════════════════ 数据加载
def load_glossary():
    """条款号 -> {en, zh, subpart_cn}"""
    p = os.path.join(ROOT, "FAR23_条款中文标题表.csv")
    out = {}
    if not os.path.exists(p):
        return out
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        out[r["条款号"]] = {
            "en": r["英文标题"], "zh": r["中文标题"],
            "sub_cn": r["所属分部（中文）"], "sub_en": r["所属分部（英文）"],
        }
    return out


def load_doc_meta():
    """(条款, 修订号, 文件类型) -> [(编号, 签发日, FR引证)]"""
    out = defaultdict(list)
    p = os.path.join(ROOT, "FAR23_条款-规则文件对应表.csv")
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        s = r["Section"].replace("Sec. ", "").strip()
        a = (r["Amendment"] or "").strip()
        t = (r["文件类型"] or "").strip()
        n = (r["文件编号"] or "").strip()
        if not n:
            continue
        out[(s, a, t)].append((n, (r["签发日期"] or "").strip(),
                               (r["FR引证"] or "").strip()))
    return out


def load_correction_hits():
    """条款号 -> [{fr, date, kind, note, file}]"""
    out = defaultdict(list)
    led = os.path.join(ROOT, "FAR23_更正通告权威台账.csv")
    meta = {}
    if os.path.exists(led):
        for r in csv.DictReader(open(led, encoding="utf-8-sig")):
            meta[r["FR引证"]] = r
    p = os.path.join(ROOT, "FAR23_受影响条款清单.csv")
    if os.path.exists(p):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            sec = (r["条款/附录"] or "").strip()
            for fr in [x.strip() for x in (r["更正件FR引证"] or "").split(";") if x.strip()]:
                m = meta.get(fr, {})
                out[sec].append({
                    "fr": fr,
                    "date": m.get("出版日", ""),
                    "kind": m.get("性质", "更正通告"),
                    "note": (r.get("更正内容要点") or m.get("说明") or "")[:400],
                    "file": m.get("本地文件", ""),
                })
    return out


# ══════════════════════════════════════════════════════ 英文 → 中文
# 修订指令高度程式化，先用规则翻译（零成本、无歧义），未命中再走翻译层。
AMEND_RULES = [
    (re.compile(r"^Section\s+23\.(\d+)\s+is\s+amended\s+by\s+revising\s+paragraph\s*\(([^)]*)\)\s*"
                r"(?:\s*and\s+paragraph\s*\(([^)]*)\))?\s*to\s+read\s+as\s+follows:?\s*$", re.I),
     lambda m: (f"§23.{m.group(1)} 条修订如下：修订 ({m.group(2)}) 段"
                + (f"与 ({m.group(3)}) 段" if m.group(3) else "") + "，内容如下：")),
    (re.compile(r"^Section\s+23\.(\d+)\s*\(([^)]*)\)\s+is\s+revised\s+to\s+read\s+as\s+follows:?\s*$", re.I),
     lambda m: f"§23.{m.group(1)} 条修订如下：修订 ({m.group(2)}) 段，内容如下："),
    (re.compile(r"^Section\s+23\.(\d+)\s+is\s+amended\s+by\s+adding\s+(?:a\s+)?new\s+paragraph\s*"
                r"\(([^)]*)\)\s*to\s+read\s+as\s+follows:?\s*$", re.I),
     lambda m: f"§23.{m.group(1)} 条新增 ({m.group(2)}) 段，内容如下："),
    (re.compile(r"^Section\s+23\.(\d+)\s+is\s+amended\s+by\s+removing\s+paragraph\s*\(([^)]*)\)\s*$", re.I),
     lambda m: f"§23.{m.group(1)} 条删除 ({m.group(2)}) 段。"),
    (re.compile(r"^Section\s+23\.(\d+)\s+is\s+amended\s*$", re.I),
     lambda m: f"§23.{m.group(1)} 条予以修订。"),
    (re.compile(r"^(?:Part\s+23\s+is\s+amended\s+by\s+)?adding\s+a\s+new\s+Sec(?:tion)?\.?\s+23\.(\d+)"
                r"(?:\s+(?:after|before)\s+Sec(?:tion)?\.?\s+23\.(\d+))?\s+to\s+read\s+as\s+follows:?\s*$", re.I),
     lambda m: (f"第 23 部新增 §23.{m.group(1)}"
                + (f"（置于 §23.{m.group(2)} 之后）" if m.group(2) else "") + "，内容如下：")),
    (re.compile(r"^by\s+amending\s+(?:the\s+)?(\w+)\s+sentence\s+of\s+Sec(?:tion)?\.?\s+23\.(\d+)"
                r"\s*(\([^)]*\))?\s*by\s+deleting\s+the\s+words?\s+[\"“]([^\"”]*)[\"”]\s*"
                r"and\s+inserting\s+the\s+words?\s+[\"“]([^\"”]*)[\"”]\s*in\s+their\s+place\.?\s*$", re.I),
     lambda m: (f"修订 §23.{m.group(2)}{m.group(3) or ''} 段的"
                f"{ {'first':'第一','second':'第二','third':'第三','fourth':'第四'}.get((m.group(1) or '').lower(), m.group(1)) }句："
                f"删除 “{m.group(4)}”，替换为 “{m.group(5)}”。")),
    (re.compile(r"^by\s+revising\s+(?:paragraph|Sec(?:tion)?\.?)\s*", re.I),
     lambda m: "修订相应段落。"),
]

ORDINAL = {"first": "第一", "second": "第二", "third": "第三", "fourth": "第四",
           "fifth": "第五", "last": "最后"}


def zh_amend(text):
    """把一条修订指令译成中文；命中规则直接返回，否则返回 None 交给翻译层。"""
    t = " ".join(text.split())
    t = re.sub(r"^\(?\d{1,2}[\.\)]\s*", "", t)      # 去指令序号 "8. "
    for rx, fn in AMEND_RULES:
        m = rx.match(t)
        if m:
            try:
                return fn(m)
            except Exception:
                break
    return None


def zh_date(eff):
    """'02/01/1965' -> '1965-02-01'"""
    if not eff:
        return "—"
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", eff.strip())
    if m:
        return f"{m.group(3)}-{m.group(1)}-{m.group(2)}"
    return eff


def year_of(eff):
    d = zh_date(eff)
    return d[:4] if d[:4].isdigit() else "?"


# ══════════════════════════════════════════════════════ 条文解析
STAR = re.compile(r"^\s*[\*\u2022·\-–—\s]{3,}$")
LABEL = re.compile(r"^\(?\s*(\d{1,2}|[ivxIVX]{1,4}|[a-z])\s*\)")


def parse_clause(lines):
    """从「read as follows」之后的行里解析 [(段落号, 文本)]。"""
    out = []
    cur_lab, cur_txt = None, []
    for raw in lines:
        s = " ".join(raw.split())
        if not s:
            continue
        if STAR.match(s):
            continue
        if re.match(r"^(?:Sec(?:tion)?\.?|§)\s*23\.\d+", s) and len(s) < 90:
            continue          # 条款抬头行
        if re.match(r"^(?:Subpart|Appendix)\s+[A-Z]", s):
            break
        m = LABEL.match(s)
        if m and (len(s) - m.end()) > 0:
            if cur_lab is not None:
                out.append((cur_lab, " ".join(cur_txt).strip()))
            cur_lab = f"({m.group(1)})"
            cur_txt = [s[m.end():].strip()]
        elif cur_lab is not None:
            cur_txt.append(s)
    if cur_lab is not None:
        out.append((cur_lab, " ".join(cur_txt).strip()))
    return [(k, v) for k, v in out if v]


RAS = re.compile(r"to\s+read\s+as\s+follows", re.I)


def version_clause_text(amend_blocks, init_text, is_initial):
    """返回本版条文 [(段落号, 文本)]；无重印全文时返回空列表。"""
    if is_initial and init_text:
        return parse_clause(init_text)
    best = []
    for blk in amend_blocks:
        joined = [l for l in blk]
        hit = -1
        for i, l in enumerate(joined):
            if RAS.search(l):
                hit = i
                break
        if hit < 0:
            continue
        body = joined[hit:]
        body[0] = RAS.split(body[0])[-1]
        got = parse_clause(body)
        if len(got) > len(best):
            best = got
    return best


def diff_status(prev, cur):
    """对比两版条文，返回 {段落号: 状态}。状态：新增 / 修订 / —"""
    pm = dict(prev)
    cm = dict(cur)
    st = {}
    for k, v in cm.items():
        if k not in pm:
            st[k] = "新增"
        elif pm[k].strip() != v.strip():
            st[k] = "修订"
        else:
            st[k] = "—"
    for k in pm:
        if k not in cm:
            st[k] = "删除"
    return st


# ══════════════════════════════════════════════════════ 取证
def collect(sec, meta, di, corr, docmeta, glos, corrhits, tr, do_tr,
            max_discuss=2):
    num = sec.replace("Sec. ", "").strip()
    g = glos.get(num, {})
    out = {
        "num": num,
        "title_en": meta["title"],
        "title_cn": g.get("zh") or meta["title"],
        "sub_cn": g.get("sub_cn", ""),
        "versions": [], "refs": OrderedDict(), "init": None,
        "corrs": corrhits.get(num, []),
    }
    prev_clause = []
    for idx, (a, v) in enumerate(meta["versions"].items()):
        nums = corr.get((sec, a), {})
        frs, nps = nums.get("Final Rule", []), nums.get("NPRM", [])
        A = S.resolve_docs(";".join(frs), "FinalRule", di, num) if frs else []
        N = S.resolve_docs(";".join(nps), "NPRM", di, num) if nps else []
        entry = {
            "amdt": a, "amdt_cn": amdt_cn(a), "eff": zh_date(v["eff"]),
            "drs": v["drs_id"], "fr_actions": v["fr_actions"],
            "fr_nums": frs, "nprm_nums": nps,
            "fr_meta": docmeta.get((num, a, "Final Rule"), []),
            "nprm_meta": docmeta.get((num, a, "NPRM"), []),
            "fr_docs": A, "nprm_docs": N,
            "amend": [], "discuss": [], "amend_cn": [], "discuss_cn": [],
            "is_initial": idx == 0,
        }
        for d in A:
            m = S.mine_doc(d, num)
            entry["amend"] += [("Final Rule", d, b) for b in m["amend"]
                               if is_amend_instruction(b[0])]
            entry["discuss"] += [("Final Rule", d, b) for b in m["discuss"][:max_discuss]]
            out["refs"][d.name] = d
        for d in N:
            m = S.mine_doc(d, num)
            entry["amend"] += [("NPRM（提案）", d, b) for b in m["amend"]
                               if is_amend_instruction(b[0])]
            entry["discuss"] += [("NPRM（提案）", d, b) for b in m["discuss"][:max_discuss]]
            out["refs"][d.name] = d

        entry["amend"] = _dedupe(entry["amend"], _amend_key)
        entry["discuss"] = _dedupe(entry["discuss"])

        # ---- 条文（英文）+ 状态
        clause = version_clause_text([b for _, _, b in entry["amend"]],
                                     None, entry["is_initial"])
        if entry["is_initial"] and not clause:
            clause = version_clause_text([], S.initial_clause_text(num), True)
        entry["clause"] = clause
        entry["status"] = diff_status(prev_clause, clause) if clause else {}
        if clause:
            prev_clause = clause

        # ---- 中文变更摘要
        blocks = [b for _, _, b in entry["amend"]]
        if blocks:
            entry["summary_cn"] = S.summarize_change(blocks[0])
        elif entry["is_initial"]:
            entry["summary_cn"] = "初始颁行，确立本条条文"
        else:
            entry["summary_cn"] = "条文未改动（本版为交叉引用或其他条款修订所致）"
        out["versions"].append(entry)

    out["init"] = S.initial_clause_text(num)

    # ---- 翻译：整条款一次性批量并发（逐块串行实测只有 2 条/分钟）
    if do_tr and tr:
        jobs = []                       # (版本号, 类别, 槽位, 待译原文)
        for vi, e in enumerate(out["versions"]):
            for ai, (src, d, blk) in enumerate(e["amend"]):
                head = " ".join(blk[0].split())
                z = zh_amend(head)      # 规则能译就不调 API
                e["amend_cn"].append(z or "")
                if not z:
                    jobs.append((vi, "a", ai, head))
            for di, (src, d, blk) in enumerate(e["discuss"]):
                e["discuss_cn"].append("")
                jobs.append((vi, "d", di, "\n".join(blk)))
        if jobs:
            res = tr.translate_many([j[3] for j in jobs])
            for (vi, kind, idx, _), r in zip(jobs, res):
                if kind == "a":
                    out["versions"][vi]["amend_cn"][idx] = r
                else:
                    out["versions"][vi]["discuss_cn"][idx] = r
    return out


def _key(blk):
    s = re.sub(r"\s+", "", "\n".join(blk))
    return re.sub(r"^\d{1,3}[\.\)]", "", s)[:160]


# 真正的修订指令行：行首（允许 "N." 序号前缀）必须是指令句式。
# mine_doc 偶尔会把讨论段（如 "In the NPRM, … Sec. 23.302(a) is revised to add …"）
# 误判成 amend，讨论文字中间恰好含 "is revised" 也会中招，所以必须锚定行首。
INSTRUCTION_RE = re.compile(
    r"^(?:[\*•·\s]*)"                       # 星号/空白
    r"(?:\(?\d{1,3}[\.\),]?\s*)?"           # 条目号："8. " / "(8) " / "17, " / "4 "
    r"(?:"
    r"(?:Section|Sec\.?)\s*23\.\d+[a-z]?(?:\s*\([^)]*\))*\s+"   # 可含多个 (a)(1)
    r"(?:is|are|would\s+be|will\s+be)\s+"
    r"(?:amended|revised|added|removed|deleted|redesignated|corrected|republished)"
    r"|Part\s+23\s+is\s+(?:amended|revised)"
    r"|By\s+(?:amending|revising|removing|deleting|adding|redesignating|correcting|striking)"
    r"|(?:A\s+new|Adding|Revising|Removing|Amending|Correcting)\s+"
    r"(?:Sec|section|a\s+new|the\s+introductory|the\s+heading|the\s+section\s+heading|"
    r"paragraph|the\s+entry)"
    r")", re.I)

# 锚定语法查不全 NPRM 的假定式写法（"Footnote 4 in Sec. 23.397(b) would be amended…"、
# "Subparagraphs (1)… would be redesigned…"），用「假定式动作词 / read as follows」补召回
_PASSIVE_VERB = re.compile(
    r"(?:would\s+be|will\s+be|is|are)\s+"
    r"(?:amended|added|revised|redesignated|deleted|removed|corrected|republished)", re.I)


def is_amend_instruction(line):
    """判断一行是否为真正的修订指令（而非讨论段里顺带出现的 "is revised"）。

    讨论段典型："…, and, for clarity, Sec. 23.302(a) is revised to add subpart D…"
    动作词出现在句中很靠后的位置，用 60 字符的位置阈值把它和指令行分开。
    """
    if INSTRUCTION_RE.match(line) or "read as follows" in line.lower():
        return True
    m = _PASSIVE_VERB.search(line)
    return bool(m and m.start() < 60)


# Part 23 全部初始版同源：1964-12-18 整部重编（CAR Part 3 → 14 CFR Part 23）
INITIAL_FR = "29 FR 17945"


def _amend_key(blk):
    """修订指令的语义键：只看指令行（blk[0]），把 NPRM 与 FR 的措辞差异归一。

    典型同义句式：
      "2. Section 23.301 is amended by revising paragraph (d) to read as follows:"（NPRM）
      "2. Section 23.301(d) is revised to read as follows:"（FR）
    → 归一为 rev23.301(d)。FR 块常带重印条文（多行），NPRM 块只有指令行，
    所以不能用全文做键，否则同一指令去不掉重。
    """
    head = " ".join(blk[0].split())
    t = re.sub(r"^\(?\d{1,2}[\.\)]\s*", "", head).strip().lower()
    t = t.replace("section", "sec").replace("paragraph", "para")
    t = t.replace("removing", "deleting")       # FR 用 removing，NPRM 用 deleting
    m = re.match(r"^sec\.?\s*23\.(\d+)\s+is\s+amended\s+by\s+revising\s+para\s*\(([^)]*)\)", t)
    if m:
        t = f"rev23.{m.group(1)}({m.group(2)})"
    else:
        m = re.match(r"^sec\.?\s*23\.(\d+)\s*\(([^)]*)\)\s+is\s+revised", t)
        if m:
            t = f"rev23.{m.group(1)}({m.group(2)})"
        else:
            m = re.match(r"^(?:part\s*23\s*is\s*amended\s*by\s*)?adding\s+a\s*new\s*sec\.?\s*23\.(\d+)", t)
            if m:
                t = f"add23.{m.group(1)}"
    t = re.sub(r"[^a-z0-9§().]", "", t)
    return t[:120]


def _dedupe(items, key_fn=_key):
    seen, out = set(), []
    for src, d, blk in items:
        k = key_fn(blk)
        if k in seen:
            for i, o in enumerate(out):
                if key_fn(o[2]) == k and o[0].startswith("NPRM") and src.startswith("Final"):
                    out[i] = (src, d, blk)
            continue
        seen.add(k)
        out.append((src, d, blk))
    return out


# ══════════════════════════════════════════════════════ 五节 + 附录
def _fr_of(doc):
    return doc.fr_citation() if doc and doc.fr else "（FR 引证未在文本中标注）"


def _match_doc(num_str, docs):
    """按编号串（Docket 4080 / Notice 64-17 / FAA-2009-0738）找到对应文献。"""
    if not num_str or not docs:
        return None
    dk, nk = S.docket_key(num_str), S.notice_key(num_str)
    for d in docs:
        if dk and d.docket and d.docket == dk:
            return d
        if nk and d.notice and d.notice == nk:
            return d
    toks = [t for t in re.split(r"[\s/]+", num_str) if len(t) >= 3]
    for d in docs:
        if any(t in d.name for t in toks):
            return d
    return None


def _meta_str(metas, docs, num, initial=False):
    """把「编号 / 签发日 / FR 引证」拼成中文串；FR 引证缺失时从文献对象补。"""
    parts = []
    for n, sign, fr in metas:
        f = fr
        if not f:
            d = _match_doc(n, docs)
            if d and d.fr:                       # 无 FR 时不用哨兵串占位
                f = d.fr_citation()
        seg = n
        if sign:
            seg += f"，{sign} 签发"
        if f:
            seg += f"，{f}"
        elif initial and "4080" in n:
            seg += f"，{INITIAL_FR}"
        parts.append(seg)
    if not parts:
        if initial:
            parts = [f"Docket 4080，{INITIAL_FR}"]
        else:
            parts = [_fr_of(d) for d in docs[:1]] or ["—"]
    return "\n".join(parts)


def sec1(doc, d, sub_cn):
    L.H(doc, f"1　§{d['num']} 修订历史汇总", 2)
    nver = len(d["versions"])
    first, last = d["versions"][0], d["versions"][-1]
    span = f"{first['eff']} → {last['eff']}"
    yrs = [int(y) for y in (year_of(first["eff"]), year_of(last["eff"])) if y.isdigit()]
    nyear = f"（约 {max(yrs) - min(yrs)} 年）" if len(yrs) == 2 else ""
    L.H(doc, "1.1　条款基本信息", 3)
    rows = [
        ["条款号", f"14 CFR § {d['num']}"],
        ["条款标题（英文原文）", d["title_en"]],
        ["中文名称", d["title_cn"]],
        ["所属分部", sub_cn or d.get("sub_cn", "")],
        ["历史版本数", f"{nver} 版（含 {first['eff']} 初始版）"],
        ["时间跨度", f"{span}　{nyear}"],
        ["初版来源", _meta_str(first["fr_meta"], first["fr_docs"], d["num"],
                                initial=True).replace("\n", "；")],
        ["当前状态", "Historical。2017 年 Amdt 23-64 性能化重构后本条体系已改写，"
                     "本卷记录 1965–2017 年间的历史沿革。"],
        ["更正通告 / CFR 编纂校正",
         "、".join(c["fr"] for c in d["corrs"]) if d["corrs"] else "无"],
    ]
    L.mk_table(doc, ["项目", "内容"], rows, widths=[4.2, 12.6], font=9)

    L.H(doc, "1.2　修订沿革总表（按时间排序）", 3)
    rows = []
    for i, v in enumerate(d["versions"], 1):
        rows.append([
            str(i), v["amdt_cn"], v["eff"],
            _meta_str(v["nprm_meta"], v["nprm_docs"], d["num"]),
            _meta_str(v["fr_meta"], v["fr_docs"], d["num"]),
            v["summary_cn"],
        ])
    L.mk_table(doc, ["#", "修订号", "生效日期", "NPRM（编号 / 签发 / FR 引证）",
                     "Final Rule（案号 / 签发 / FR 引证）", "本版对 §%s 的修订内容" % d["num"]],
               rows, widths=[0.8, 1.9, 2.0, 3.5, 3.5, 5.1], font=8.5,
               align_center_cols=(0, 1, 2))

    L.H(doc, "1.3　修订脉络一览", 3)
    chain = []
    for i, v in enumerate(d["versions"]):
        tag = f"{year_of(v['eff'])} {'初版' if v['is_initial'] else v['amdt_cn'].replace('Amdt ', '')}"
        brief = v["summary_cn"].split("；")[0][:22]
        chain.append(f"{tag}　{brief}".strip())
    L.P(doc, "　→　".join(chain), size=10.5, bold=True,
        color=RGBColor(0x1F, 0x4E, 0x79))


DRIVERS = [
    ("事故调查与研究驱动", re.compile(
        r"accident|NTSB|crash|test data|research|study|investigation|survivable", re.I)),
    ("产业豁免压力倒逼", re.compile(
        r"exemption|petition|manufacturer|economic|cost|burden|industry", re.I)),
    ("国际协调", re.compile(r"harmoniz|JAA|JAR|EASA|Transport Canada|ARAC|Europe", re.I)),
    ("新技术 / 新构型", re.compile(
        r"new technology|composite|turbojet|canard|tandem wing|special condition|"
        r"new configuration|advanced", re.I)),
    ("从处方性走向性能化", re.compile(r"performance[- ]based|performance standard|flexibility", re.I)),
    ("澄清性与编辑性修订", re.compile(r"clarif|editorial|typograph|correction|technical amendment|"
                                 r"redesignat|obsolete", re.I)),
]


def sec2(doc, d):
    L.H(doc, f"2　§{d['num']} 各阶段修订背景与原因深度剖析", 2)
    n = 0
    for v in d["versions"]:
        n += 1
        if v["is_initial"]:
            title = f"2.{n}　初始颁行（{v['eff']}，{v['amdt_cn']}）：{d['title_cn']}"
        else:
            brief = v["summary_cn"].split("；")[0][:26]
            title = f"2.{n}　{v['amdt_cn']}（{v['eff']} 生效）：{brief}"
        L.H(doc, title, 3)

        np = _meta_str(v["nprm_meta"], v["nprm_docs"], d["num"]).replace("\n", "；")
        fr = _meta_str(v["fr_meta"], v["fr_docs"], d["num"]).replace("\n", "；")
        L.BULLET(doc, "修订沿革：", [
            (f"本版由 NPRM　{np}　提出；最终规章 {fr}，{v['eff']} 生效。", "normal")])
        if v["amend_cn"]:
            L.BULLET(doc, "修订指令（中文译文）：", [("　".join(v["amend_cn"]), "normal")])
        elif v["amend"]:
            L.BULLET(doc, "修订指令（英文原文）：",
                     [("　".join(" ".join(b[0].split()) for _, _, b in v["amend"][:2]), "normal")])
        else:
            L.BULLET(doc, "本版对条文的处置：", [(v["summary_cn"], "normal")])

        if v["discuss_cn"]:
            for i, zh in enumerate(v["discuss_cn"], 1):
                L.BULLET(doc, f"背景与评论处置（{i}）：", [(zh, "normal")])
        elif v["discuss"]:
            for i, (src, doc_, blk) in enumerate(v["discuss"], 1):
                L.BULLET(doc, f"背景与评论处置（{src}，英文原文 {i}）：",
                         [("\n".join(blk)[:1800], "normal")])

        if v["amend"]:
            L.P(doc, "修订指令英文原文（供核对）：", size=9, bold=True,
                color=RGBColor(0x80, 0x80, 0x80), indent=0.6, before=4, after=1)
            for _, doc_, blk in v["amend"][:2]:
                head = " ".join(blk[0].split())
                L.P(doc, head[:700], size=8.5, indent=1.0,
                    color=RGBColor(0x60, 0x60, 0x60))

        if v["clause"] and not v["is_initial"]:
            L.BULLET(doc, "本版重印条文段落：",
                     [("、".join(k for k, _ in v["clause"]), "normal")])

    # ---- 更正件专节
    if d["corrs"]:
        n += 1
        L.H(doc, f"2.{n}　更正通告与 CFR 编纂校正", 3)
        L.P(doc, "CFR 官方溯源注通常只登记修订案的 FR 引证，更正通告多被吞并或只留一个裸页码；"
                 "CFR 编纂校正（OFR 在年度版排印时发出的 Correction）更是不进溯源注。"
                 "下列文献经官方 CFR 溯源注、federalregister.gov API 与 govinfo 影印件三源交叉核实。",
            size=10)
        rows = [[c["fr"], c["date"], c["kind"], c["note"]] for c in d["corrs"]]
        L.mk_table(doc, ["FR 引证", "日期", "性质", f"对 §{d['num']} 的实际改动"],
                   rows, widths=[2.0, 2.0, 3.2, 9.6], font=8.5)
        L.info_box(doc, "引用提示", [
            "更正件通常改动拼写、交叉引用编号或表头，不产生新的 Amdt 号，"
            "DRS 条款版本索引与 CFR 溯源注都可能漏记；引用时请以现行 CFR 文本为准。",
        ], bg=L.C_SUB_BG)


def sec3(doc, d):
    L.H(doc, f"3　§{d['num']} 条款原文演变对比", 2)
    L.P(doc, "本节先给出颜色图例，再按版本逐条列出条文（英文原文），"
             "最后以段落结构矩阵横向对照各版本。条文表的「状态」列标明"
             "该段在本版是新增、修订、删除还是继承未变。", size=10)
    L.legend(doc)
    n = 0
    for i, v in enumerate(d["versions"], 1):
        if not v["clause"]:
            continue
        n += 1
        src = _meta_str(v["fr_meta"], v["fr_docs"], d["num"],
                        initial=v["is_initial"]).replace("\n", "；")
        tag = "初始版" if v["is_initial"] else v["amdt_cn"]
        L.H(doc, f"3.{n}　版本{n}：{tag}（{v['eff']} 生效）　来源：{src}", 3)
        if v["is_initial"]:
            note = "1965 年重编版，源自 CAR Part 3 相应条款。"
        else:
            note = v["summary_cn"]
        L.P(doc, note, size=9.5, italic=True, color=L.C_DEL, indent=0.3)
        rows = []
        for lab, txt in v["clause"]:
            st = v["status"].get(lab, "—")
            seg_style = "new" if st == "新增" else ("rev" if st == "修订" else
                                                    ("del" if st == "删除" else "normal"))
            rows.append([lab, [(txt, seg_style)],
                         st if st != "—" else "—（继承上版）"])
        L.clause_table(doc, "", rows)
    if n == 0:
        L.P(doc, "（本地文献中未检索到本条各版本的重印全文，"
                 "请参照 1.2 沿革总表与各版修订指令。）", size=9.5)

    # ---- 段落结构演变矩阵
    L.H(doc, f"3.{n + 1}　段落结构演变矩阵（横向对照 {len(d['versions'])} 个版本）", 3)
    withtxt = [v for v in d["versions"] if v["clause"]]
    if not withtxt:
        L.P(doc, "（无可用全文，矩阵略。）", size=9.5)
        return
    labels = []
    for v in withtxt:
        for k, _ in v["clause"]:
            if k not in labels:
                labels.append(k)
    header = ["段落"] + [f"{v['amdt_cn']}\n({v['eff']})" for v in withtxt]
    rows = []
    for lab in labels:
        row = [lab]
        for v in withtxt:
            st = v["status"].get(lab)
            if lab not in dict(v["clause"]):
                row.append("○ 无")
            elif st == "新增":
                row.append("● 新增")
            elif st == "修订":
                row.append("△ 修订")
            else:
                row.append("● 有")
        rows.append(row)
    w = [1.4] + [max(1.9, 13.8 / max(len(withtxt), 1))] * len(withtxt)
    L.mk_table(doc, header, rows, widths=w, font=8, align_center_cols=(0,))


def sec4(doc, d):
    L.H(doc, f"4　§{d['num']} 总结", 2)
    rows = []
    for v in d["versions"]:
        blob = "\n".join("\n".join(b) for _, _, b in v["discuss"]) + "\n" + \
               "\n".join("\n".join(b) for _, _, b in v["amend"])
        if not blob.strip():
            continue
        for name, rx in DRIVERS:
            m = rx.search(blob)
            if m:
                i = max(0, m.start() - 60)
                snippet = " ".join(blob[i:m.start() + 160].split())
                rows.append([name, snippet + "…",
                             "初始版" if v["is_initial"] else v["amdt_cn"]])
                break
    if rows:
        L.P(doc, "下表按「驱动力」归纳本条各次修订的成因（依据各版 NPRM/Final Rule 原文关键词自动归纳）。",
            size=10)
        L.mk_table(doc, ["驱动力", "在本条上的体现（原文摘录）", "对应版本"],
                   rows, widths=[3.0, 10.4, 3.4], font=8.5)
    else:
        L.P(doc, "本条各版本的修订动因未在 NPRM/Final Rule 前言中展开说明，"
                 "多为随主修订案一并调整的编辑性或交叉引用性修改。", size=10)

    nver = len(d["versions"])
    amdts = ", ".join(v["amdt_cn"] for v in d["versions"])
    got = sum(1 for v in d["versions"] if v["amend"])
    L.info_box(doc, "本条取证提示", [
        f"共 {nver} 个历史版本：{amdts}。",
        f"其中 {got} 个版本在本地文献中检索到明确修订指令原文；"
        f"{sum(1 for v in d['versions'] if v['clause'])} 个版本检索到重印全文。",
        ("本条另有更正通告 / CFR 编纂校正：" + "、".join(c["fr"] for c in d["corrs"])
         + "，DRS 与 CFR 溯源注均可能漏记。") if d["corrs"]
        else "本条未见更正通告。",
    ], bg=L.C_SUB_BG)


def sec5(doc, d):
    L.H(doc, "5　参考文献", 2)
    L.P(doc, "下列文件均已下载并保存于本地目录「FAR 23/Rulemaking Docs」；"
             "FR 引证格式为「卷号 FR 起始页码」。文件名保留原名以便按名检索。", size=10)
    rows = []
    for i, (name, doc_) in enumerate(d["refs"].items(), 1):
        kind = "NPRM（提案）" if doc_.kind == "NPRM" else "Final Rule（颁行）"
        rows.append([f"[{i}]", kind, doc_.label(), _fr_of(doc_), name])
    if d["corrs"]:
        for c in d["corrs"]:
            if c.get("file"):
                rows.append([f"[{len(rows) + 1}]", "更正通告 / 校正", c["fr"],
                             c["date"], os.path.basename(c["file"])])
    if d["init"]:
        rows.append([f"[{len(rows) + 1}]", "修订记录",
                     "FAA 动态法规系统（DRS）条款修订版索引",
                     "Part 23 @ Amdt 23-63 as of 08/29/2017 and earlier",
                     "FAR23_Amdt23-63_条款清单.csv（本地）"])
        rows.append([f"[{len(rows) + 1}]", "条文校对", "1965 年重编原始规章（初始条文）",
                     "29 FR 17955，1964-12-18",
                     "FinalRule_Docket 4080_Amdt23-0_1964-12-18_0002.txt"])
    L.mk_table(doc, ["编号", "类型", "文件编号 / 案号", "Federal Register 引证与日期",
                     "本地文件名 / 来源"],
               rows, widths=[1.0, 2.4, 4.0, 4.2, 5.2], font=8)


def appendix(doc, d):
    L.H(doc, "附录　DRS 条款版本记录 ID", 2)
    L.P(doc, f"FAA 动态法规系统（DRS）中 §{d['num']} 各历史版本的记录 ID，"
             "可用于在 DRS 中直接定位相应版本条文。", size=10)
    rows = [[v["amdt_cn"], v["eff"], v["drs"] or "—"] for v in d["versions"]]
    L.mk_table(doc, ["修订号", "生效日期", "DRS 记录 ID（docUniqueId）"],
               rows, widths=[2.6, 3.0, 11.2], font=8.5)


def one_clause(doc, d, sub_cn):
    L.H(doc, f"§ {d['num']}　{d['title_cn']}（{d['title_en']}）", 1)
    L.P(doc, f"{d['versions'][0]['eff']} 初版　｜　"
             f"{d['versions'][-1]['eff']} {d['versions'][-1]['amdt_cn']}　｜　"
             f"共 {len(d['versions'])} 个历史版本",
        size=10.5, color=RGBColor(0x59, 0x59, 0x59))
    sec1(doc, d, sub_cn)
    sec2(doc, d)
    sec3(doc, d)
    sec4(doc, d)
    sec5(doc, d)
    appendix(doc, d)


# ══════════════════════════════════════════════════════ 卷装配
def new_doc():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.0)
    s.top_margin, s.bottom_margin = Cm(2.0), Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name = L.EN_FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(L.qn("w:eastAsia"), L.CN_FONT)
    st.paragraph_format.line_spacing = 1.35
    L.setup_heading_styles(doc)
    L.add_page_number_footer(s)
    return doc


def cover(doc, sub_key, data):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    L.para_spacing(p, 40, 4)
    L.set_run(p.add_run("14 CFR Part 23 条款修订历史与背景分析报告"),
              size=22, bold=True, color=RGBColor(0x1F, 0x4E, 0x79))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    L.para_spacing(p, 0, 6)
    L.set_run(p.add_run(SUBPART_CN.get(sub_key, sub_key)),
              size=17, bold=True, color=RGBColor(0x2E, 0x74, 0xB5))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    L.para_spacing(p, 0, 16)
    L.set_run(p.add_run(SUBPART_KEY.get(sub_key, sub_key)),
              size=11, color=RGBColor(0x59, 0x59, 0x59))
    nver = sum(len(d["versions"]) for d in data)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    L.para_spacing(p, 0, 18)
    L.set_run(p.add_run(f"快照：Part 23 @ Amdt 23-63 as of 08/29/2017 and earlier　｜　"
                        f"共 {len(data)} 个条款 / {nver} 个历史版本"),
              size=11, color=RGBColor(0x59, 0x59, 0x59))

    L.info_box(doc, "文档说明与编制依据", [
        [("• 修订日期、修订案号、NPRM 与 Final Rule 编号：取自 ", "normal"),
         ("FAA 动态法规系统（DRS）", "bold"),
         ("「Part 23 @ Amdt 23-63 as of 08/29/2017 and earlier」条款修订版索引。", "normal")],
        "• 修订指令与讨论/评论处置段落：逐条引自本地已下载的 NPRM 与 Final Rule 官方文本，"
        "已译为中文；条文英文原文另附以供核对。",
        "• 更正通告与 CFR 编纂校正：由官方 CFR 2016 版溯源注、federalregister.gov API 与"
        "govinfo 整期影印件三源交叉验证，详见 FAR23_更正通告权威台账.csv。",
        "• 每个条款统一按「五节 + 一附录」编排：1 修订历史汇总 / 2 各阶段修订背景与原因深度剖析 / "
        "3 条款原文演变对比 / 4 总结 / 5 参考文献 / 附录 DRS 条款版本记录 ID。",
        "• 标题已写入 Word 大纲级别（1 级＝条款，2 级＝五节与附录，3 级＝小节，4 级＝子题），"
        "可在导航窗格或大纲视图中按级跳转。",
    ])


def outline(doc, sub_key, data):
    L.H(doc, "本卷大纲", 1)
    L.P(doc, "下按条款列出本卷结构。每一条款下统一设 1–5 节与附录；"
             "二级及以下标题可在 Word 导航窗格中展开。", size=10)
    rows = []
    for d in data:
        rows.append([f"§ {d['num']}", d["title_cn"], d["title_en"],
                     str(len(d["versions"])),
                     "、".join(c["fr"] for c in d["corrs"]) or "—"])
    L.mk_table(doc, ["条款", "中文名称", "英文标题（原文）", "版本数", "更正件"],
               rows, widths=[1.8, 5.6, 6.4, 1.4, 2.4], font=8.5,
               align_center_cols=(0, 3))
    doc.add_page_break()


def build(sub_key, limit=None, out=None, do_translate=True, workers=10):
    key = SUBPART_KEY.get(sub_key, sub_key)
    cn = SUBPART_CN.get(sub_key, sub_key)
    print("建索引…", flush=True)
    di = S.build_doc_index()
    corr = S.load_correspondence()
    glos = load_glossary()
    docmeta = load_doc_meta()
    corrhits = load_correction_hits()
    secs = S.load_sections(key)
    print(f"  {len(secs)} 个条款", flush=True)

    tr = None
    if do_translate:
        try:
            import translate as T
            tr = T.Translator(workers=workers, verbose=False)
            print(f"  翻译缓存 {len(tr.cache)} 条", flush=True)
        except Exception as e:
            print(f"  !! 翻译层不可用（{e}），讨论段落将保留英文", flush=True)
            tr = None

    data = []
    for i, (sec, meta) in enumerate(secs.items()):
        if limit and i >= limit:
            break
        d = collect(sec, meta, di, corr, docmeta, glos, corrhits, tr, do_translate)
        data.append(d)
        print(f"  [{i + 1}/{len(secs) if not limit else limit}] §{d['num']}　"
              f"{d['title_cn']}　{len(d['versions'])} 版", flush=True)
        if tr:
            tr.save()

    doc = new_doc()
    cover(doc, sub_key, data)
    outline(doc, sub_key, data)
    L.H(doc, "目录", 1)
    L.add_toc_field(doc, "1-2")
    doc.add_page_break()
    for d in data:
        one_clause(doc, d, cn)

    out = out or os.path.join(ROOT, f"FAR23_Subpart{sub_key}_修订史.docx")
    doc.save(out)
    print("已生成:", out)
    print(f"   条款 {len(data)}　段落 {len(doc.paragraphs)}　表格 {len(doc.tables)}")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("subpart", nargs="?", default="C")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--out", default=None)
    ap.add_argument("--no-translate", action="store_true")
    ap.add_argument("--workers", type=int, default=10)
    a = ap.parse_args()
    build(a.subpart, a.limit or None, a.out,
          do_translate=not a.no_translate, workers=a.workers)
