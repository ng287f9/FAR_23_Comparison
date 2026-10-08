#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""分部卷严格校核（比 check_volume.py 严得多）。

校核维度：
  A 骨架   每条款 6 个 H2 且标题逐一匹配模板；H1/H2/H3 数量关系；H3 内部小节齐全
  B 语言   除「条款原文 / 英文原文块 / 来源行 / 条款标题括注」外，任何段落中文占比
           低于阈值即判为疑似未译（按 H2 所属节做白名单，不靠关键词猜）
  C 内容   每版本有「修订沿革：」；每条款第 3 节有版本小节与演变矩阵；讨论段与来源标注配平
  D 异常   空段、空表、None/nan/\\ufffd/TODO/未标注引证 等占位串
  E 一致性 与 DRS 源数据的条款数、版本数比对

用法：
    python audit_volume.py <docx> [更多 docx …]
    python audit_volume.py --all            # 校核项目内全部 *_详版.docx
    python audit_volume.py --all --json out.json
"""
import argparse
import collections
import glob
import json
import os
import re
import sys

from docx import Document

ROOT = os.path.dirname(os.path.abspath(__file__))
CJK = re.compile(r"[\u4e00-\u9fff]")

# 每条款应有的 6 个 H2（用正则，条款号位置通配）
SEC6 = [
    re.compile(r"^1\u3000§[\w.]+\s*修订历史汇总$"),
    re.compile(r"^2\u3000§[\w.]+\s*各阶段修订背景与原因深度剖析$"),
    re.compile(r"^3\u3000§[\w.]+\s*条款原文演变对比$"),
    re.compile(r"^4\u3000§[\w.]+\s*总结$"),
    re.compile(r"^5\u3000参考文献$"),
    re.compile(r"^附录\u3000DRS 条款版本记录 ID$"),
]
SEC6_CN = ["修订历史汇总", "各阶段修订背景与原因深度剖析", "条款原文演变对比",
           "总结", "参考文献", "附录 DRS 记录 ID"]

# 卷首固定 H1
FRONT_H1 = ["本卷大纲", "目录"]

BAD_PAT = re.compile(
    r"(?:\bNone\b|\bnan\b|\bNaN\b|\\ufffd|\ufffd|\bTODO\b|\bFIXME\b|"
    r"\bundefined\b|\bnull\b|\[未找到\]|（FR 引证未在文本中标注）|待补)")

ZH_MIN = 0.15          # 普通段落的中文占比下限

# 这些前缀的段落天然含大量英文（编号清单、文献名、英文原文块），豁免语言检查
EXEMPT_PREFIX = (
    "▪ 修订沿革", "▪ 本版重印条文段落", "▪ 本版对条文的处置", "▪ 修订指令（英文原文）",
    "▪ 条文结构", "▪ 更正", "快照：", "Subpart ", "Appendix ",
)
EN_BLOCK_LABELS = ("英文原文：", "修订指令英文原文（供核对）：")

# 修订指令的中文译文行：几乎不含中文（真未译）才判缺陷。
# 阈值放宽是因为已译行里 §/条款号/编号占了大半篇幅。
ZH_AMEND_MIN = 0.08


def zh_ratio(s):
    return len(CJK.findall(s)) / max(len(s), 1)


def para_kind(p, state):
    """给段落打标签，决定是否豁免语言检查。"""
    t = p.text.strip()
    st = p.style.name
    if not t:
        return "empty"
    if st == "Heading 1":
        return "h1"
    if st == "Heading 2":
        return "h2"
    if st == "Heading 3":
        return "h3"
    if state["in_en_block"]:
        return "en_original"
    if t == "英文原文：" or t.startswith("英文原文："):
        return "en_label"
    if t.startswith("来源："):
        return "source"
    if state["sec_idx"] == 3:          # 第 3 节整节都是英文条文
        return "clause_text"
    if t.startswith("■") or t.startswith("▪") and "版本" in t:
        return "legend"
    return "normal"


def audit(path):
    r = {
        "file": os.path.basename(path), "ok": True,
        "err": [], "warn": [], "stat": {},
    }
    d = Document(path)
    paras = d.paragraphs

    h1, h2, h3 = [], [], []
    cur_clause = None
    cur_sec = 0                    # 当前条款内第几节（1..6）；0=不在条款内
    sec_of_clause = collections.OrderedDict()
    clause_of_h2 = []
    in_en = False
    state = {"sec_idx": 0, "in_en_block": False}

    untranslated, anomalies, empty_paras = [], [], 0
    amend_untranslated = []
    discuss = 0
    source_lines = 0
    amend_lines = 0
    in_outline = False

    for p in paras:
        t = p.text.strip()
        st = p.style.name
        if not t:
            empty_paras += 1
            continue
        # ---- 标题推进状态
        if st == "Heading 1":
            h1.append(t)
            cur_clause, cur_sec, in_en = None, 0, False
            # 卷首「本卷大纲」到第一个条款标题之间的大纲区豁免语言检查
            in_outline = (t in FRONT_H1) or in_outline
            state = {"sec_idx": 0, "in_en_block": False, "in_outline": in_outline}
            m = re.match(r"^§\s*([\w.]+)\u3000(.+)$", t)
            if m:
                cur_clause = m.group(1)
                in_outline = False
                state["in_outline"] = False
                sec_of_clause[cur_clause] = []
            continue
        state["in_outline"] = in_outline
        if st == "Heading 2":
            h2.append(t)
            if cur_clause:
                cur_sec += 1
                sec_of_clause[cur_clause].append(t)
                clause_of_h2.append((cur_clause, cur_sec, t))
                state["sec_idx"] = cur_sec
                state["in_en_block"] = False
                in_en = False
            continue
        if st == "Heading 3":
            h3.append(t)
            state["in_en_block"] = False
            in_en = False
            continue
        # ---- 正文
        # 英文原文块：连续小字段落，遇到 ▪ / 来源 / 标题结束
        if state["in_en_block"]:
            if t.startswith("▪") or t.startswith("来源：") or t.startswith("■"):
                state["in_en_block"] = False
            else:
                continue
        if any(t.startswith(x) for x in EN_BLOCK_LABELS):
            state["in_en_block"] = True
            continue
        if t.startswith("来源："):
            source_lines += 1
            continue
        # 修订指令中文译文行：中文占比过低 = 指令未译（真缺陷，单独计）
        if t.startswith("▪ 修订指令（中文译文）："):
            body = t.split("：", 1)[1] if "：" in t else ""
            cjk = len(CJK.findall(body))
            # 极短行（如"§23.615 被删除。"）中文本来就少，不能只看占比
            if cjk == 0 or (cjk < 4 and len(body) > 25) or (
                    len(body) > 25 and zh_ratio(body) < ZH_AMEND_MIN):
                amend_untranslated.append(t[:130])
            continue
        if t.startswith("▪ 背景与评论处置"):
            discuss += 1
        if t.startswith("▪ 修订沿革"):
            amend_lines += 1
        # 豁免：编号清单 / 文献名 / 大纲 / 快照 / 条款标题
        if any(t.startswith(x) for x in EXEMPT_PREFIX):
            continue
        if state["sec_idx"] == 3:          # 第 3 节整节是英文条文
            continue
        if state["in_outline"]:
            continue
        if BAD_PAT.search(t):
            anomalies.append(t[:110])
        if zh_ratio(t) < ZH_MIN:
            untranslated.append((state["sec_idx"], t[:130]))

    # ---------- A 骨架 ----------
    n_clause = len(sec_of_clause)
    bad6 = {k: len(v) for k, v in sec_of_clause.items() if len(v) != 6}
    # 六节标题逐一匹配
    title_mismatch = []
    for cl, secs in sec_of_clause.items():
        for i, s in enumerate(secs):
            if i < len(SEC6) and not SEC6[i].match(s):
                title_mismatch.append(f"{cl} 第{i+1}节标题异常：{s[:60]}")
    # 第 1 / 3 节 H3 小节
    h3_by_clause = collections.OrderedDict()
    cur = None
    for p in paras:
        t, st = p.text.strip(), p.style.name
        if st == "Heading 1":
            m = re.match(r"^§\s*([\w.]+)\u3000", t)
            cur = m.group(1) if m else None
            if cur:
                h3_by_clause.setdefault(cur, [])
        elif st == "Heading 3" and cur:
            h3_by_clause[cur].append(t)
    missing_h3 = []
    for cl, hs in h3_by_clause.items():
        for need in ("1.1", "1.2", "1.3"):
            if not any(x.startswith(need) for x in hs):
                missing_h3.append(f"{cl} 缺 {need}")
        if not any("段落结构演变矩阵" in x for x in hs):
            missing_h3.append(f"{cl} 缺 演变矩阵")

    expect_h2 = n_clause * 6
    expect_h1 = n_clause + len(FRONT_H1)

    # ---------- C 内容 ----------
    # 每条款至少 1 个版本小节：H3 "3.x 版本" 计数
    ver_sections = sum(1 for x in h3 if re.match(r"^3\.\d+\u3000版本", x))
    # 第 2 节小节标题的质量瑕疵：残留 "3. " 条目号，或以 "：" 结尾（指令被截断）
    title_blemish = []
    for x in h3:
        if not re.match(r"^2\.\d+\u3000", x):
            continue
        head = x.split("\u3000", 1)[1]
        if re.match(r"^\d{1,3}\s*[\.\)、]\s", head):
            title_blemish.append(f"{x[:70]}（残留条目号）")
        elif head.rstrip().endswith("："):
            title_blemish.append(f"{x[:70]}（尾部冒号，指令被截断）")
    # 版本总数：以「修订沿革」行计（每版一行）
    versions_total = sum(1 for p in paras
                         if p.text.strip().startswith("▪ 修订沿革"))
    # 每条款 1.2 表格 + 附录表格
    tables = d.tables
    empty_tables = sum(1 for tb in tables if len(tb.rows) == 0 or len(tb.columns) == 0)

    r["stat"] = {
        "条款": n_clause, "H1": len(h1), "H2": len(h2), "H3": len(h3),
        "段落": len(paras), "表格": len(tables), "空表": empty_tables,
        "版本小节": ver_sections, "讨论段": discuss, "来源标注": source_lines,
        "修订沿革行": amend_lines, "空段": empty_paras,
        "疑似未译": len(untranslated), "指令未译": len(amend_untranslated),
        "异常串": len(anomalies), "标题瑕疵": len(title_blemish),
        "无条文版本": versions_total - ver_sections,
    }

    if len(h1) != expect_h1:
        r["err"].append(f"H1 数 {len(h1)} ≠ 条款数+2 = {expect_h1}")
    if len(h2) != expect_h2:
        r["err"].append(f"H2 数 {len(h2)} ≠ 条款数×6 = {expect_h2}")
    if bad6:
        r["err"].append(f"节数不为 6 的条款 {len(bad6)} 个：" +
                        ", ".join(f"{k}={v}" for k, v in list(bad6.items())[:6]))
    if title_mismatch:
        r["err"] += title_mismatch[:5]
    if missing_h3:
        r["err"] += missing_h3[:8]
    if amend_untranslated:
        r["err"].append(f"修订指令未译 {len(amend_untranslated)} 条")
    if untranslated:
        r["err"].append(f"疑似未译 {len(untranslated)} 段")
    if anomalies:
        r["err"].append(f"异常占位串 {len(anomalies)} 处")
    if empty_tables:
        r["err"].append(f"空表格 {empty_tables} 个")
    if title_blemish:
        r["err"].append(f"第 2 节标题残留序号/尾部冒号 {len(title_blemish)} 处")
    if discuss != source_lines:
        r["err"].append(f"讨论段 {discuss} ≠ 来源标注 {source_lines}")
    if versions_total > ver_sections:
        r["warn"].append(
            f"有 {versions_total - ver_sections} 个版本无「条文原文」小节"
            f"（共 {versions_total} 版，写了 {ver_sections} 版）")

    r["ok"] = not r["err"]
    r["untranslated"] = [x[1] for x in untranslated[:8]]
    r["amend_untranslated"] = amend_untranslated[:8]
    r["title_blemish"] = title_blemish[:8]
    r["anomalies"] = anomalies[:8]
    return r


def cross_check_source(results):
    """与 DRS 源数据比对条款数、版本数。"""
    try:
        import part23_source as S
    except Exception as e:
        return [f"跳过源数据比对（{e}）"]
    msgs = []
    key_of = {}
    import build_clause_doc as B
    for k, v in B.SUBPART_KEY.items():
        key_of[k] = v
    for r in results:
        name = r["file"]
        m = re.match(r"FAR23_Subpart(.+?)_条款修订历史与背景分析_详版\.docx$", name)
        if not m:
            continue
        sub = m.group(1)
        key = key_of.get(sub)
        if not key:
            continue
        try:
            secs = S.load_sections(key)
        except Exception as e:
            msgs.append(f"{sub}: 源数据读取失败 {e}")
            continue
        doc_n = r["stat"]["条款"]
        if doc_n != len(secs):
            msgs.append(f"{sub}: 文档条款 {doc_n} ≠ DRS {len(secs)}")
        src_ver = sum(len(v["versions"]) for v in secs.values())
        if r["stat"]["版本小节"] < src_ver:
            msgs.append(f"{sub}: 文档版本小节 {r['stat']['版本小节']} < DRS 版本数 {src_ver}")
    return msgs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--json")
    a = ap.parse_args()

    files = a.files
    if a.all or not files:
        files = sorted(glob.glob(os.path.join(ROOT, "FAR23_Subpart*_详版.docx")))

    results = [audit(f) for f in files]
    order = ["A", "B", "C", "D", "E", "F", "G",
             "AppA", "AppB", "AppC", "AppD", "AppE", "AppF",
             "AppG", "AppH", "AppI", "AppJ"]
    def k(r):
        m = re.match(r"FAR23_Subpart(.+?)_", r["file"])
        s = m.group(1) if m else r["file"]
        return (order.index(s) if s in order else 99, s)
    results.sort(key=k)

    for r in results:
        s = r["stat"]
        flag = "合格" if r["ok"] else f"问题 {len(r['err'])} 项"
        print(f"【{flag}】{r['file']}")
        print(f"    条款 {s['条款']}  段落 {s['段落']}  表格 {s['表格']}"
              f"（空 {s['空表']}）  H1={s['H1']} H2={s['H2']} H3={s['H3']}")
        print(f"    版本 {s['修订沿革行']}（有原文 {s['版本小节']}）  讨论段 {s['讨论段']}"
              f"  来源标注 {s['来源标注']}  空段 {s['空段']}"
              f"  未译 {s['疑似未译']}  指令未译 {s['指令未译']}")
        for e in r["err"]:
            print(f"    !! {e}")
        for w in r["warn"]:
            print(f"    ~  {w}")
        for t in r["untranslated"]:
            print(f"      未译示例：{t}")
        for t in r.get("amend_untranslated", []):
            print(f"      指令未译：{t}")
        for t in r.get("title_blemish", []):
            print(f"      标题瑕疵：{t}")
        for t in r["anomalies"]:
            print(f"      异常示例：{t}")

    src = cross_check_source(results)
    if src:
        print("\n【源数据比对】")
        for m in src:
            print("  !!", m)
    else:
        print("\n【源数据比对】条款数与版本数全部一致")

    bad = [r for r in results if not r["ok"]]
    print(f"\n合计 {len(results)} 卷，合格 {len(results)-len(bad)}，有问题 {len(bad)}")
    if a.json:
        json.dump(results, open(a.json, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("明细已写:", a.json)
    return 1 if (bad or src) else 0


if __name__ == "__main__":
    sys.exit(main())
