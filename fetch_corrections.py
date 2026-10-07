#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T2 — 全量扫描 1994–2017 年 FAA 在 Federal Register 上涉及 14 CFR Part 23 的
     Rule 类文档，按 action 区分：

       correction      : action 含 "correct"（更正通告，修正已发布规章的错误）
       amendment       : action 含 "technical amendment"（技术性修订，无 Amdt 号的独立 Rule）
       final rule      : action 为 "Final rule"（正式最终规章，带 Amdt 号）
       other           : 其余

     再按 23.xxx 条款号、与官方 CFR 2016 溯源注里「裸 FR 引证」清单做交叉比对：
       - 命中溯源注 -> 标记为「溯源注引证」
       - 未命中      -> 标记为「溯源注未列（DRS/CFR 漏记）」

输出：Rulemaking Docs/Corrections/ 下的 .txt / _meta.json
      根目录 FAR23_1994后Part23规则文件总表.csv
"""
import csv
import gzip
import json
import re
import time
import urllib.parse
import urllib.request
from collections import OrderedDict
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Rulemaking Docs" / "Corrections"
API = "https://www.federalregister.gov/api/v1/documents.json"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

FIELDS = ("document_number", "title", "type", "action", "start_page", "end_page",
          "publication_date", "agencies", "docket_ids", "citation",
          "raw_text_url", "html_url", "pdf_url")

# Part 23 相关判定模式
RE_AMD23 = re.compile(r"\b23[-\s]??\d{1,2}\b")            # Amdt 23-54
RE_SEC23 = re.compile(r"\b23\.\d{1,4}\b")                  # 条款 23.573
RE_PART23 = re.compile(r"part\s*23\b", re.I)


def _get(url, timeout=120, retry=3):
    last = None
    for _ in range(retry):
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": UA, "Accept-Encoding": "gzip"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                raw = r.read()
                if r.headers.get("Content-Encoding") == "gzip":
                    raw = gzip.decompress(raw)
                return raw.decode("utf-8", "replace")
        except Exception as e:
            last = e
            time.sleep(2)
    raise last


def _fetch_window(gte, lte):
    """拉一个时间窗内的 FAA Rule 文档（含分页，规避 API 10000 条上限）"""
    params = [
        ("per_page", "1000"),
        ("order", "oldest"),
        ("conditions[agencies][]", "federal-aviation-administration"),
        ("conditions[type][]", "RULE"),
        ("conditions[publication_date][gte]", gte),
        ("conditions[publication_date][lte]", lte),
    ]
    for f in FIELDS:
        params.append(("fields[]", f))
    base = API + "?" + urllib.parse.urlencode(params)
    got, page = [], 1
    while True:
        data = json.loads(_get(f"{base}&page={page}"))
        res = data.get("results") or []
        got.extend(res)
        total = data.get("count") or 0
        if total > 10000:
            raise RuntimeError(f"窗口 {gte}..{lte} 有 {total} 条 > 10000，需再细分")
        if not res or len(got) >= total:
            break
        page += 1
        time.sleep(0.3)
    return got


def fetch_pages(year_from=1994, year_to=2017):
    """用 CFR 精确过滤（Title 14 Part 23）拉 Rule 类文档。

    federalregister.gov 支持 conditions[cfr][title]/[part]，能把结果从上万条压到几百条，
    既快又不会撞上 API 的 10000 条结果上限。若某年窗口仍超限则自动细分。
    """
    base_params = [
        ("per_page", "1000"),
        ("order", "oldest"),
        ("conditions[cfr][title]", "14"),
        ("conditions[cfr][part]", "23"),
        ("conditions[type][]", "RULE"),
        ("conditions[publication_date][gte]", f"{year_from}-01-01"),
        ("conditions[publication_date][lte]", f"{year_to}-12-31"),
    ]
    for f in FIELDS:
        base_params.append(("fields[]", f))
    base = API + "?" + urllib.parse.urlencode(base_params)
    docs, page = [], 1
    while True:
        data = json.loads(_get(f"{base}&page={page}"))
        res = data.get("results") or []
        docs.extend(res)
        print(f"    page {page}  累计 {len(docs)} / {data.get('count')}")
        if not res or len(docs) >= (data.get("count") or 0):
            break
        page += 1
        time.sleep(0.3)
    return docs


def is_part23(d):
    """判断该文档是否与 14 CFR Part 23 有关"""
    blob = " ".join(str(d.get(k) or "") for k in ("title", "docket_ids", "action"))
    dockets = " ".join(d.get("docket_ids") or [])
    # docket 里出现 Amdt 23-xx 是最强证据
    if re.search(r"amendment\s+no\.?\s*[^.]*\b23\s*[-–]\s*\d", dockets, re.I):
        return True, "docket-amdt23"
    if RE_PART23.search(blob):
        return True, "title-part23"
    if re.search(r"\b23\.\d{2,4}\b", blob):
        return True, "title-sec23"
    return False, ""


def classify(d):
    """按 FR 的 action 字段 + 文档号形态细分文档性质。

    cfr-correction     : CFR Correction —— OFR 在 CFR 年度版编纂时的排版/重复段落校正，
                         发文机构是 Office of the Federal Register 而非 FAA，
                         标题固定为 "Airworthiness Standards: ... CFR Correction"
    doc-correction     : 文档号带 C 前缀（如 C1-2016-28714），FR 官方的「更正某文档」标记
    correction         : FAA 更正通告，action/标题含 correction
    technical-amendment: 技术性修订，无实质变更的措辞/交叉引用更新
    final-rule         : 正式最终规章（带 Amdt 号）
    special-conditions : 专用条件（个案机型）
    """
    act = (d.get("action") or "").lower()
    title = (d.get("title") or "").lower()
    num = d.get("document_number") or ""
    if "cfr correction" in title:
        return "cfr-correction"
    if re.match(r"^C\d+[-_]", num):
        return "doc-correction"
    if "special condition" in title or "special condition" in act:
        return "special-conditions"
    if re.search(r"correct", act) or re.search(r"\bcorrection\b", title):
        return "correction"
    if "technical amendment" in act:
        return "technical-amendment"
    if "final rule" in act:
        return "final-rule"
    if "affirmation" in act:
        return "affirmation"
    return act.strip() or "(none)"


# 抽取被改动条款的正则：覆盖 "§ 23.573"、"Sec. 23.561"、"in Sec. 23.1511" 诸写法
RE_SEC_ANY = re.compile(r"(?:§|sec\.|section|\bsections\b)\s*23\.\d{1,4}", re.I)
RE_SEC_BARE = re.compile(r"23\.\d{2,4}")


def affected_sections(body: str):
    """从正文抽取被真正改动的 Part 23 条款号。

    优先取 amendatory instruction 区（PART 23 之后）或 CFR Correction 语句本身。
    """
    if not body:
        return []
    cut = max(body.find("PART 23"), body.find("CFR Correction"),
              body.find("List of Subjects"))
    seg = body[cut:] if cut > 0 else body
    nums = set()
    for m in RE_SEC_ANY.finditer(seg):
        nums.add(re.search(r"23\.\d+", m.group(0)).group(0))
    if not nums:
        for m in RE_SEC_BARE.finditer(seg):
            nums.add(m.group(0))
    return sorted(nums, key=lambda s: [int(p) for p in s.split(".")])


def load_source_note_hits():
    """从官方 CFR 溯源注 CSV 取出被判为『裸 FR 引证』的 1994 年后清单"""
    p = ROOT / "FAR23_更正通告清单_官方口径.csv"
    if not p.exists():
        return {}
    out = {}
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        fr = (r.get("更正件FR引证") or "").strip()
        m = re.match(r"^(\d{1,3})\s+FR\s+(\d{1,6})$", fr)
        if not m:
            continue
        y = int((r.get("更正日期") or "")[:4] or 0)
        if y >= 1994:
            out[f"{m.group(1)} FR {m.group(2)}"] = r
    return out


def main(year_from=1994, year_to=2017, do_download=True):
    OUT.mkdir(parents=True, exist_ok=True)
    print(f"扫描 {year_from}-{year_to} 年 FAA Rule 文档 ...")
    docs = fetch_pages(year_from, year_to)

    # CFR part=23 已由 API 精确过滤，此处不再二次筛，
    # 只用 is_part23 标注命中方式，便于人工复核漏网。
    hits = docs
    for d in hits:
        ok, how = is_part23(d)
        d["_match"] = how or "cfr-filter-only"
        d["_kind"] = classify(d)

    print(f"\n命中 {len(hits)} 篇")
    by_kind = {}
    for d in hits:
        by_kind.setdefault(d["_kind"], []).append(d)
    for k, v in sorted(by_kind.items(), key=lambda kv: -len(kv[1])):
        print(f"    {k:<22} {len(v)}")

    src = load_source_note_hits()
    print(f"\n官方 CFR 溯源注里 1994 年后的『裸 FR 引证』{len(src)} 条: {sorted(src)}")
    for k in sorted(src):
        mark = "✓已识别" if any((d.get("citation") or "").strip() == k for d in hits) else "✗未命中"
        print(f"      {k:<14} {mark}")

    # ---- 写总表 ----
    out_csv = ROOT / "FAR23_1994后Part23规则文件总表.csv"
    rows = []
    for d in sorted(hits, key=lambda x: x["publication_date"]):
        cite = (d.get("citation") or "").strip()
        rows.append(OrderedDict([
            ("FR引证", cite),
            ("出版日", d["publication_date"]),
            ("文档号", d["document_number"]),
            ("性质", d["_kind"]),
            ("action", (d.get("action") or "").strip()),
            ("标题", (d.get("title") or "").strip()),
            ("Docket", "; ".join(d.get("docket_ids") or [])),
            ("起页", d.get("start_page")),
            ("止页", d.get("end_page")),
            ("溯源注是否引证", "是" if cite in src else "否"),
            ("溯源注影响条款", (src.get(cite, {}).get("受影响条款") or "")),
            ("正文提及条款", ""),
            ("本地文件", ""),
        ]))

    def write_rows():
        with open(out_csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    write_rows()
    print(f"→ {out_csv.name}  ({len(rows)} 行)")

    # ---- 下载全部「非普通最终规章/专用条件」的文档正文 ----
    keep = {"correction", "cfr-correction", "doc-correction",
            "technical-amendment", "(none)"}
    targets = [d for d in hits if d["_kind"] in keep]
    print(f"\n待处理正文 {len(targets)} 篇")
    for d in sorted(targets, key=lambda x: x["publication_date"]):
        num = d["document_number"]
        date = d["publication_date"]
        y, mth, dd = date.split("-")
        vol_match = re.match(r"^(\d{1,3}) FR (\d{1,6})$", d.get("citation") or "")
        frv, frp = vol_match.groups() if vol_match else ("??", "??????")
        KIND_TAG = {"correction": "Correction",
                    "cfr-correction": "CFRCorrection",
                    "doc-correction": "Correction",
                    "technical-amendment": "TechAmend",
                    "(none)": "Unclassified"}
        stem = f"{KIND_TAG.get(d['_kind'], 'Rule')}_{frv}FR{frp}_{date}_{num}"

        txt_p = OUT / f"{stem}.txt"
        meta_p = OUT / f"{stem}_meta.json"
        meta_p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")

        if txt_p.exists():
            status = f"SKIP(已有 {txt_p.stat().st_size:,}B)"
        elif not do_download:
            status = "未下载(--no-download)"
        else:
            status = "FAIL"
            for url in (f"https://www.federalregister.gov/documents/full_text/text/"
                        f"{y}/{mth}/{dd}/{num}.txt",
                        d.get("raw_text_url") or ""):
                if not url:
                    continue
                try:
                    t = _get(url)
                    if t.strip():
                        txt_p.write_text(t, encoding="utf-8")
                        status = f"OK {len(t):,}c"
                        break
                except Exception as e:
                    status = f"FAIL {e}"
                time.sleep(0.5)

        # ---- 下载后二次定性：CFR Correction 只能从正文识别（API 标题里没有）----
        secs = ""
        if txt_p.exists():
            body = txt_p.read_text(encoding="utf-8", errors="replace")
            if re.search(r"CFR\s+Correction", body, re.I) and d["_kind"] == "(none)":
                d["_kind"] = "cfr-correction"
                new_p = OUT / f"CFRCorrection_{frv}FR{frp}_{date}_{num}.txt"
                txt_p.rename(new_p)
                txt_p = new_p
                stem = txt_p.stem
                (OUT / f"{stem}_meta.json").write_text(
                    json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
                # 清掉 Unclassified 版元数据
                (OUT / f"Unclassified_{frv}FR{frp}_{date}_{num}_meta.json").unlink(missing_ok=True)
            secs = ", ".join(affected_sections(body))

        print(f"    {stem}\n         {status}  [{d['_kind']}] {(d.get('action') or '')[:40]}"
              + (f"\n         影响条款: {secs}" if secs else ""))

        for r in rows:
            if r["文档号"] == num and r["出版日"] == date:
                r["性质"] = d["_kind"]
                if txt_p.exists():
                    r["本地文件"] = f"Corrections/{txt_p.name}"
                if secs:
                    r["正文提及条款"] = secs

    write_rows()
    print(f"→ {out_csv.name}  已刷新本地文件列")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="y0", type=int, default=1994)
    ap.add_argument("--to", dest="y1", type=int, default=2017)
    ap.add_argument("--no-download", action="store_true")
    main(ap.parse_args().y0, ap.parse_args().y1, not ap.parse_args().no_download)
