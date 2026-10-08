#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 23 条款修订史 —— 数据源层。

职责：
  1. 读 DRS 条款清单 / 条款-规则文件对应表
  2. 扫描 Rulemaking Docs 建立"文献索引"（文件名 + 内嵌 DRS 头 → 类型/编号/FR 引证/日期）
  3. 按 (条款, 修订号) 解析出对应的 NPRM / Final Rule 本地文件
  4. 从文献正文里挖掘该条款的：NPRM 提案原文、NPRM 说明、FR 评论处置、规章修订文本
  5. 从 1964 年初版规章（Docket 4080）里抽取条款初始全文

全部为本地文件操作，不联网。
"""
import csv
import json
import os
import re
from collections import defaultdict, OrderedDict

ROOT = "/Users/glennchou/FAR 23"
DOCS = os.path.join(ROOT, "Rulemaking Docs")
CSV_SECTIONS = os.path.join(ROOT, "FAR23_Amdt23-63_条款清单.csv")
CSV_CORR = os.path.join(ROOT, "FAR23_条款-规则文件对应表.csv")
CSV_DL = os.path.join(ROOT, "FAR23_已下载规则文件清单.csv")

INITIAL_DOC = "FinalRule_Docket 4080_Amdt23-0_1964-12-18_0002.txt"


# ---------------------------------------------------------------- 基础工具
def _norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def compact_paragraphs(path):
    """读文件 → 紧凑段落列表（CR/LF 通吃，去掉空行与多余空白）。"""
    try:
        raw = open(path, encoding="utf-8", errors="replace").read()
    except OSError:
        return []
    if path.lower().endswith(".pdf"):
        try:
            import fitz
            raw = "\n".join(p.get_text() for p in fitz.open(path))
        except Exception:
            return []
    else:
        raw = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", raw)
        raw = re.sub(r"(?s)<[^>]+>", "\n", raw)
        raw = (raw.replace("&nbsp;", " ").replace("&amp;", "&")
                  .replace("&lt;", "<").replace("&gt;", ">")
                  .replace("&quot;", '"').replace("&#39;", "'"))
    lines = [_norm(l) for l in raw.split("\n")]
    return [l for l in lines if l]


def docket_key(s):
    m = re.search(r"Docket\s*(?:No\.?|Number)?\s*:?\s*(FAA-[\d-]+|\d[\d,]*)", s or "", re.I)
    if m:
        return re.sub(r"[,\s]", "", m.group(1))
    m = re.search(r"\b(FAA-20\d\d-\d+)\b", s or "")
    return m.group(1) if m else None


def notice_key(s):
    m = re.search(r"Notice\s*(?:No\.?)?\s*:?\s*(\d{2}-\d{1,2})", s or "", re.I)
    return m.group(1) if m else None


def fr_key(s):
    """从字符串里提取 (卷, 起始页)。注意文件名里是 _34FR13078_ 形式，前面是下划线，
    不能用 \\b（下划线是词字符），故用 (?<![\\w])。"""
    s = s or ""
    m = re.search(r"(?<!\d)(\d{2,3})\s*FR\s*(\d{1,6})(?!\d)", s, re.I)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    # 只有卷号没有页码（如 52FR_1987-01-15）
    m = re.search(r"(?<!\d)(\d{2,3})\s*FR(?!\d)", s, re.I)
    if m:
        return (int(m.group(1)), 0)
    return None


def date_key(s):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s or "")
    return m.group(0) if m else None


# ---------------------------------------------------------------- 文献索引
class Doc:
    __slots__ = ("path", "name", "kind", "docket", "notice", "fr", "date",
                 "paras", "_loaded", "alts")

    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path)
        self.kind = "FinalRule" if self.name.startswith("FinalRule") else (
            "NPRM" if self.name.startswith("NPRM") else "Other")
        self.docket = docket_key(self.name)
        self.notice = notice_key(self.name)
        self.fr = fr_key(self.name)
        self.date = date_key(self.name)
        self.paras = None
        self._loaded = False
        self.alts = []

    def text(self):
        if not self._loaded:
            self.paras = compact_paragraphs(self.path)
            self._loaded = True
        return self.paras

    # ---- 内嵌 DRS 头（比文件名更权威，优先用）----
    def enrich(self):
        ps = self.text()
        head = ps[:80]
        for i, p in enumerate(head):
            low = p.lower()
            if low.startswith("docket number") or low.startswith("docket:"):
                d = docket_key(p)
                if d:
                    self.docket = d
            elif low.startswith("citation"):
                f = fr_key(p)
                d = date_key(p)
                if f:
                    self.fr = f
                else:
                    # 形如 "[Federal Register: April 11, 2008 (Volume 73, Number 69)]"
                    mv = re.search(r"Volume\s+(\d+)", p, re.I)
                    if mv:
                        self.fr = (int(mv.group(1)), self.fr[1] if self.fr else 0)
                if d:
                    self.date = d
            elif low.startswith("page number"):
                pg = re.search(r"(\d{3,6})", p)
                if pg and self.fr:
                    self.fr = (self.fr[0], int(pg.group(1)))
            elif low.startswith("cfr nprm") or low.startswith("final rules"):
                pass
        # notice 号只在文件名里出现
        if not self.notice:
            self.notice = notice_key(self.name)
        return self

    def label(self):
        """人类可读的文献标识"""
        if self.notice:
            return f"Notice {self.notice} / Docket {self.docket or '—'}"
        if self.docket:
            return f"Docket {self.docket}"
        return self.name

    def fr_citation(self):
        return f"{self.fr[0]} FR {self.fr[1]}" if self.fr else "（FR 引证未在文本中标注）"


def build_doc_index():
    idx = []
    for fn in sorted(os.listdir(DOCS)):
        if fn.startswith(".") or fn == "_meta":
            continue
        p = os.path.join(DOCS, fn)
        if not os.path.isfile(p):
            continue
        if fn.lower().endswith(".json"):
            continue
        # 同名 .txt 优先于 .html（.txt 是从 .html 提取的，等价但更小）
        idx.append(Doc(p))
    # 去重：同一文献的 txt/html/pdf 都保留，但按 (kind,docket,notice,fr) 分组装文
    groups = defaultdict(list)
    for d in idx:
        if d.kind == "Other":
            continue
        groups[(d.kind, d.docket, d.notice, d.fr)].append(d)
    out = OrderedDict()
    for k, v in groups.items():
        # 首选 txt，其次 html，最后 pdf
        def rank(d):
            e = d.name.lower()
            return 0 if e.endswith(".txt") else (1 if e.endswith(".html") else 2)
        v.sort(key=rank)
        rep = v[0].enrich()
        rep.alts = v[1:]
        out[k] = rep
    return out


# ---------------------------------------------------------------- 条款数据
def amdt_sort_key(a):
    a = (a or "").strip()
    if a.lower() in ("initial", "unnamed", ""):
        return (0, "")
    m = re.match(r"(\d+)-(\d+)([A-Z]?)", a)
    if m:
        return (int(m.group(1)) * 1000 + int(m.group(2)), m.group(3))
    m = re.match(r"(\d+)", a)
    return (int(m.group(1)) if m else 0, "")


def load_sections(subpart=None):
    """→ OrderedDict[section] = {amdt: {...}}（按修订号时间序）"""
    rows = list(csv.DictReader(open(CSV_SECTIONS, encoding="utf-8-sig")))
    if subpart:
        rows = [r for r in rows if r["Subpart/Appendix"] == subpart]
    secs = defaultdict(dict)
    titles = {}
    for r in rows:
        s = _norm(r["Section"])
        titles[s] = _norm(r["Title"])
        a = _norm(r["Amendment"])
        secs[s][a] = {
            "amdt": a,
            "eff": _norm(r["Effective Date"]),
            "fr_actions": _norm(r["Final Rule Actions"]),
            "drs_id": _norm(r["DRS Doc ID"]),
        }
    out = OrderedDict()
    for s in sorted(secs, key=lambda x: [int(t) if t.isdigit() else 0
                                         for t in re.findall(r"\d+", x)] or [0]):
        d = secs[s]
        out[s] = {
            "title": titles[s],
            "versions": OrderedDict(sorted(d.items(), key=lambda kv: amdt_sort_key(kv[0]))),
        }
    return out


def load_correspondence():
    """→ {(section, amdt): {'NPRM': [...编号], 'Final Rule': [...编号]}}"""
    rows = list(csv.DictReader(open(CSV_CORR, encoding="utf-8-sig")))
    out = defaultdict(lambda: defaultdict(list))
    for r in rows:
        s = _norm(r["Section"])
        a = _norm(r["Amendment"])
        t = _norm(r["文件类型"])
        n = _norm(r["文件编号"])
        if not n:
            continue
        # 一条编号里可能含多个 docket
        for num in re.split(r"[;]", n):
            num = _norm(num)
            if num and num not in out[(s, a)][t]:
                out[(s, a)][t].append(num)
    return out


# ---------------------------------------------------------------- 文献解析
def resolve_docs(num, kind, doc_index, secnum=None):
    """把对应表里的编号字符串解析到本地文献对象，返回按相关度排序的列表。

    同一 Docket 可能对应多份文献（如 Docket 27805 既有 1996 年 61 FR 5138 正文，
    又有 2008 年技术更正件）。排序规则：
      1) 文献里真的提到该条款的排前面（命中数多者优先）
      2) 其次取 FR 卷页小的（即时间更早的正文）
    """
    # Doc.kind 的取值是 "FinalRule"/"NPRM"（无空格）。历史调用方有传 "Final Rule"
    # （带空格）的，导致 `k == kind` 永不成立、退化成「不区分类型」——NPRM 会被误当
    # Final Rule 引用。这里统一去空格，彻底免疫此类拼写差异。
    kind = (kind or "").replace(" ", "")
    dk = docket_key(num)
    nk = notice_key(num)
    cands = []
    seen = set()
    # 先按 notice 号（NPRM 老式编号）
    if nk:
        for (k, d, n, f), doc in doc_index.items():
            if k == kind and n == nk and id(doc) not in seen:
                cands.append(doc)
                seen.add(id(doc))
    # 再按 docket 号
    if dk:
        for (k, d, n, f), doc in doc_index.items():
            if k == kind and d and d.startswith(dk) and id(doc) not in seen:
                cands.append(doc)
                seen.add(id(doc))
    if not cands and dk:  # 退而求其次：不区分 kind
        for (k, d, n, f), doc in doc_index.items():
            if d and d.startswith(dk) and id(doc) not in seen:
                cands.append(doc)
                seen.add(id(doc))
    if secnum and len(cands) > 1:
        def score(doc):
            ps = doc.text()
            rx = sec_regex(secnum)
            hits = sum(1 for p in ps if rx.search(p))
            return (-hits, doc.fr or (999, 999))
        cands.sort(key=score)
    else:
        cands.sort(key=lambda d: (d.fr or (999, 999)))
    return cands


# ---------------------------------------------------------------- 正文挖掘
AMEND_VERB = re.compile(
    r"\b(?:amend(?:ed|ing|s)?|revise[sd]?|revising|redesignat\w*|remov\w*|"
    r"delet\w*|add(?:ed|ing|s)?|rescind\w*)\b", re.I)
# 只认"货真价实"的修订指令句式，避免把讨论段落误判成指令
AMEND_TAIL = re.compile(
    r"(?:\bis amended\b|\bamended by\b|\bby amending\b|\bby revising\b|\bby adding\b|"
    r"\bby removing\b|\bby deleting\b|\bis revised\b|\brevised by\b|\bis revised to read\b|"
    r"\bis added to read\b|\ba new (?:Sec(?:tion)?\.?|§)[\s\d.]+is added\b|\bis added\b|"
    r"\bis removed\b|\bis redesignated\b|\bto read as follows\b|\bread as follows\b|"
    r"\bwould be amended\b|\bwould be added\b|\bwould be revised\b)", re.I)
# 讨论段落的开头标志——命中即判为讨论，优先于修订指令
DISCUSS_OPEN = re.compile(
    r"^(?:the notice proposed|it was proposed|the faa proposed|we proposed|"
    r"one commenter|two commenter|three commenter|several comment|a commenter|"
    r"comments? (?:were|was) received|the commenter|commenters? |comment to|"
    r"this proposal|proposal \d|in response|discussion|explanation|the proposal|"
    r"we propose|the faa (?:agrees|disagrees|has |considers|concluded|does not|is not|"
    r"understands|intends|determined|received)|no comments|the notice also|"
    r"portions of|section \d+\.\d+ would be)", re.I)
DISCUSS_MARK = re.compile(
    r"\b(?:comment(?:er|s|ed)?|FAA (?:agrees|disagrees|does not agree|has determined|"
    r"considers?|concluded?|proposes?|intends?|understands?)|we (?:agree|disagree|propose)|"
    r"this proposal|proposal \d|notice no\.|petition|recommend|concurs?|"
    r"the commenter|commenters?|is the same as|in the NPRM|as proposed|"
    r"no comments|would be revised|would be amended|would be added|"
    r"harmoniz\w+|clarif\w+|for clarification|technical amendment)\b", re.I)
NEW_ITEM = re.compile(r"^\s*\d{1,3}\s*[.\)]\s*(?:By\s+)?(?:[Aa]mend|[Ss]ec(?:tion)?\.?\s*23\.|§\s*23\.)")
NEW_ITEM2 = re.compile(r"^\s*\d{1,3}\s*[.\)]\s+[A-Z]")
SECTION_HEAD = re.compile(r"^(?:Sec(?:tion)?\.?|§)\s*23\.\d+")
# 附录修订指令专用句式（附录条款号在正文里不出现，只能靠这句识别）。
# 实测出现过的写法：
#   "20. By adding a new Appendix G to Part 23 to read as follows:"
#   "78. Part 23 is amended by adding a new appendix H to read as follows:"
#   "Appendix D is amended by revising paragraph (b) to read as follows:"
APPENDIX_INSTR_RE = re.compile(
    r"^(?:[\*•·\s]*)(?:\(?\d{1,3}[\.\),]?\s*)?"
    r"(?:"
    r"By\s+(?:adding|amending|revising|removing|deleting|redesignating|striking|correcting)\b"
    r"|Part\s+23\s+is\s+(?:amended|revised)\s+by\s+"
    r"(?:adding|amending|revising|removing|deleting|redesignating|striking|correcting)\b"
    r"|(?:adding|amending|revising|removing|deleting|redesignating|striking)\s+"
    r"(?:a\s+new\s+)?[Aa]ppendix\s+[A-J]\b"
    r"|(?:A\s+new\s+)?[Aa]ppendix\s+[A-J]\b[^.]{0,60}?\bis\s+"
    r"(?:amended|revised|added|removed|redesignated|deleted)\b"
    r")", re.I)


def sec_regex(secnum):
    """'23.561' → 匹配 'Sec. 23.561' / 'Section 23.561' / '§ 23.561' 但避免 23.5611

    附录条款（A23.1 / G23.3 …）在 FR 正文里通常裸写，不带 "Sec." 前缀
    （如 "39. Appendix A is amended by revising … section A23.1, paragraphs A23.11(c)(1)…"），
    故对字母开头的编号额外允许裸写形式。
    """
    body = secnum.replace(".", r"\.")
    if re.match(r"^[A-Za-z]", secnum):
        return re.compile(r"(?:(?:Sec(?:tion)?\.?|§)\s*|(?<![A-Za-z0-9]))" + body + r"(?!\d)")
    return re.compile(r"(?:Sec(?:tion)?\.?|§)\s*" + body + r"(?!\d)")


def appendix_alias(secnum):
    """附录条款的「字母别名」正则。

    FR 前言论述附录时几乎从不写条款号（不写 "D23.1"），只写 "appendix D"。
    只用条款号去挖会漏掉附录的全部背景论述，故额外提供这一别名。
    返回 None 表示非附录条款（不走别名）。
    """
    m = re.match(r"^([A-Z])\d", secnum or "")
    if not m:
        return None
    return re.compile(r"[Aa]ppendix\s+" + m.group(1) + r"\b(?!\w)")


def find_hits(paras, secnum, max_hits=40, alias=None):
    """返回 [(idx, kind)]，kind ∈ {'amend','discuss'}

    alias 为附录字母别名正则（见 appendix_alias）：附录条款号在 FR 正文里不出现，
    必须靠别名 + 附录专用指令句式才能定位修订指令与重印全文。
    """
    rx = sec_regex(secnum)

    def _is_part23(p):
        """只认第 23 部的修订指令，避免把 "…to Part 25 to read as follows" 收进来。"""
        m = re.search(r"\bPart\s+(\d+)", p)
        return (m is None) or (m.group(1) == "23")

    hits = []
    for i, p in enumerate(paras):
        by_alias = bool(alias and alias.search(p))
        # 长段落一般不参与（多是整段重印，会把讨论与正文混在一起），
        # 但附录修订指令有时就把整篇重印并进同一段，故指令句式命中时破例放行。
        instr = bool(by_alias and APPENDIX_INSTR_RE.match(p) and _is_part23(p))
        if len(p) > 3000 and not instr:
            continue
        by_sec = rx.search(p)
        if not (by_sec or by_alias):
            continue
        # 排除"其他条款正文里的交叉引用"：以段落符号/括号开头且是纯法条
        stripped = p.lstrip()
        is_reg_text = bool(re.match(r"^\([a-z0-9ivx]+\)", stripped)) or \
            bool(re.match(r"^\[?[a-z]\)\s", stripped))
        if by_sec:
            # 条款号命中：保持原有判定顺序（前言句 > 修订指令 > 论述标记）
            if DISCUSS_OPEN.search(p):
                hits.append((i, "discuss"))
            elif AMEND_TAIL.search(p) and AMEND_VERB.search(p) and not is_reg_text:
                hits.append((i, "amend"))
            elif DISCUSS_MARK.search(p) and not is_reg_text:
                hits.append((i, "discuss"))
        else:
            # 只靠附录别名命中：先认「附录修订指令」的专用句式，其余才算前言论述。
            # 不能一律按论述处理——附录的指令写作
            # "20. By adding a new Appendix G to Part 23 to read as follows:"，
            # 段落里并不出现条款号 G23.1，否则重印全文就挖不到。
            if instr or (APPENDIX_INSTR_RE.match(p) and _is_part23(p)):
                hits.append((i, "amend"))
            elif DISCUSS_OPEN.search(p):
                hits.append((i, "discuss"))
            elif DISCUSS_MARK.search(p) and not is_reg_text:
                hits.append((i, "discuss"))
        if len(hits) >= max_hits:
            break
    return hits


def grab_amendment(paras, start, secnum, max_paras=30, alias=None):
    """从修订指令段开始，抓取整段修订文本（到下一个修订条目为止）。

    alias（附录字母别名）参与「是否本条款」的判断：附录重印全文里的
    "Appendix G--…" 小节标题不含条款号，只认条款号会误判为下一条而提前截断。
    """
    rx = sec_regex(secnum)

    def mine(p):
        return bool(rx.search(p) or (alias and alias.search(p)))

    out = [paras[start]]
    for j in range(start + 1, min(start + max_paras, len(paras))):
        p = paras[j]
        if NEW_ITEM.match(p) or (NEW_ITEM2.match(p) and not mine(p)):
            break
        if SECTION_HEAD.match(p) and out and not rx.search(p):
            break
        out.append(p)
        if re.fullmatch(r"[*\s•\-]{3,}", p) and len(out) > 3:
            # 星号分隔后再多看 2 段，若还是星号/新条目则停
            nxt = paras[j + 1] if j + 1 < len(paras) else ""
            if NEW_ITEM.match(nxt) or NEW_ITEM2.match(nxt):
                break
    # 去掉尾部连续的星号行
    while out and re.fullmatch(r"[*\s•\-]{3,}", out[-1]):
        out.pop()
    return out


def grab_discussion(paras, start, secnum, window=3, max_paras=8):
    """讨论段 + 其后续紧邻的相关段落（如 'Explanation.'）。"""
    out = [paras[start]]
    for j in range(start + 1, min(start + window, len(paras))):
        p = paras[j]
        if SECTION_HEAD.match(p) or NEW_ITEM.match(p) or NEW_ITEM2.match(p):
            break
        if re.match(r"^(?:Explanation|Discussion)\b", p) or \
           re.match(r"^(?:The |This |In response|FAA )", p):
            out.append(p)
        else:
            break
        if len(out) >= max_paras:
            break
    return out


def mine_doc(doc, secnum, want=("amend", "discuss"), alias=None):
    """从一篇文献里挖出与该条款有关的全部摘录。

    alias 给附录条款用（见 appendix_alias）：FR 前言只写 "appendix D" 不写 "D23.1"。
    """
    paras = doc.text()
    if not paras:
        return {"amend": [], "discuss": []}
    res = {"amend": [], "discuss": []}
    seen = set()
    for i, kind in find_hits(paras, secnum, alias=alias):
        if kind not in want:
            continue
        if kind == "amend":
            blk = grab_amendment(paras, i, secnum,
                                 max_paras=(400 if alias else 30), alias=alias)
        else:
            blk = grab_discussion(paras, i, secnum)
        key = (kind, i)
        if key in seen:
            continue
        # amend 块去重（同一条修订指令可能被多次命中）
        # 注意：res[kind] 里存的是 (i, blk)，取首段要用 b[1][0]
        sig = blk[0][:120]
        if any(b[1][0][:120] == sig for b in res[kind]):
            continue
        seen.add(key)
        res[kind].append((i, blk))
    # 修订指令块排序：以编号开头（如 "40. Section 23.561 is amended…"）的优先，
    # 其次含 "to read as follows" 的，最后按出现位置
    def akey(item):
        i, blk = item
        head = blk[0]
        starts_num = 0 if re.match(r"^\s*\d{1,3}\s*[\.\)]\s", head) else 1
        has_trf = 0 if re.search(r"to read as follows|read as follows", head, re.I) else 1
        return (starts_num, has_trf, i)
    res["amend"] = [b for _, b in sorted(res["amend"], key=akey)]
    res["discuss"] = [b for _, b in sorted(res["discuss"], key=lambda x: x[0])]
    return res


# ---------------------------------------------------------------- 初始条文
def initial_clause_text(secnum):
    """从 1964 年 Docket 4080 的全文中抽取该条款初始版全文。"""
    p = os.path.join(DOCS, INITIAL_DOC)
    paras = compact_paragraphs(p)
    rx = re.compile(r"^Sec\.\s*" + secnum.replace(".", r"\.") + r"\b")
    start = None
    for i, x in enumerate(paras):
        if rx.match(x):
            start = i
            break
    if start is None:
        return None
    out = [paras[start]]
    for j in range(start + 1, min(start + 60, len(paras))):
        x = paras[j]
        if re.match(r"^Sec\.\s*23\.\d+\b", x):
            break
        if re.match(r"^(?:Subpart|Appendix)\s+[A-Z]", x):
            break
        out.append(x)
    while out and re.fullmatch(r"[*\s•\-]{3,}", out[-1]):
        out.pop()
    return out


# ---------------------------------------------------------------- 变更摘要
def summarize_change(amend_block):
    """从修订指令里自动提炼一句中文变更摘要。"""
    if not amend_block:
        return "（未见明确修订指令）"
    txt = " ".join(amend_block)
    low = txt.lower()
    acts = []
    for pat, zh in (
        (r"by adding (?:a )?new paragraph", "新增段落"),
        (r"adding new paragraph", "新增段落"),
        (r"by adding (?:new )?(?:sub)?paragraph", "新增段落"),
        (r"\badding\b", "新增内容"),
        (r"by revising paragraph", "修订段落"),
        (r"revising paragraph", "修订段落"),
        (r"\brevising\b|\brevis", "修订"),
        (r"by removing paragraph|removing paragraph|is removed", "删除段落"),
        (r"\bremoving\b|\bdeleting\b|\brescinding\b", "删除内容"),
        (r"redesignat", "重新编号"),
        (r"is amended by", "修订"),
        (r"is added to read", "新增条款"),
    ):
        if re.search(pat, low):
            acts.append(zh)
            break
    # 抓出被改动的段落号
    paras = re.findall(r"paragraphs?\s*\(([^)]{1,80})\)", txt)
    paras = [f"({p.strip()})" for p in paras]
    dup = []
    for p in paras:
        if p not in dup:
            dup.append(p)
    scope = "、".join(dup[:6])
    act = acts[0] if acts else "修订"
    return f"{act}" + (f"：{scope}" if scope else "")


if __name__ == "__main__":
    idx = build_doc_index()
    print("文献数:", len(idx))
    n_fr = sum(1 for d in idx.values() if d.kind == "FinalRule")
    n_np = sum(1 for d in idx.values() if d.kind == "NPRM")
    print(f"  FinalRule {n_fr} / NPRM {n_np}")
    secs = load_sections("Subpart C - Structure")
    print("Subpart C 条款:", len(secs))
    corr = load_correspondence()
    ok = tot = 0
    for s in list(secs)[:10]:
        for a in secs[s]["versions"]:
            for t in ("Final Rule", "NPRM"):
                for num in corr.get((s, a), {}).get(t, []):
                    tot += 1
                    if resolve_doc(num, "FinalRule" if t == "Final Rule" else "NPRM", idx):
                        ok += 1
    print(f"前10条款文献解析成功率: {ok}/{tot}")
