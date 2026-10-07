#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T2 补完 — 1994 年前更正件只能从 govinfo 整期影印 PDF 按页抽取

流程：
  1. 下载 https://www.govinfo.gov/content/pkg/FR-<日期>/pdf/FR-<日期>.pdf
  2. 在整篇里定位 FR 页码数字 target_page 所在的 PDF 页
     （FR 每个印刷页自带页码，用 x 投影 + 独立小块检测最稳）
  3. 用 fr_column_extract 做分栏重排，落到 Corrections/

用法：
    python fetch_old_corrections.py            # 按 TARGETS 全部处理
    python fetch_old_corrections.py 52 FR 34745  # 只处理一条
"""
import gzip
import json
import re
import shutil
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CORR = ROOT / "Rulemaking Docs" / "Corrections"
TMP = Path("/tmp/fr_whole_issue")
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

# (FR卷, FR页码, 出版日, 影响条款, 预估整期体积)
TARGETS = [
    (34, 14727, "1969-09-24", "Appendix A", 6),
    (32, 13714, "1967-09-30", "23.1325", 11),
    (34, 17509, "1969-10-30", "23.369", 15),
    (32, 13505, "1967-09-27", "23.1325", 15),
    (35, 1102,  "1970-01-28", "SFAR No. 23", 19),
    (38, 32784, "1973-11-28", "23.155", 24),
    (52, 34745, "1987-09-14", "23.3/23.1199/23.1323/23.1351/App F/App G", 38),
    (52, 7262,  "1987-03-09", "23.1201/23.443", 44),
    (56, 5455,  "1991-02-11", "23.161/23.423/23.701", 50),
    (53, 34194, "1988-09-02", "23.807/23.811", 67),
    (55, 46888, "1990-11-07", "23.1321", 161),
    (58, 27060, "1993-05-06", "23.1091/23.1191/23.1305/23.961/23.971", 301),
    (58, 18975, "1993-04-09", "23.1193", 481),
]


def download(date_iso, max_mb):
    """下载整期 PDF 到 /tmp/fr_whole_issue/，返回路径"""
    TMP.mkdir(parents=True, exist_ok=True)
    dst = TMP / f"FR-{date_iso}.pdf"
    if dst.exists() and dst.stat().st_size > 1024:
        return dst
    url = f"https://www.govinfo.gov/content/pkg/FR-{date_iso}/pdf/FR-{date_iso}.pdf"
    print(f"    下载整期 {date_iso}（约 {max_mb} MB）...")
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=600) as r, open(dst, "wb") as f:
        shutil.copyfileobj(r, f)
    print(f"    -> {dst.stat().st_size/1048576:.1f} MB")
    return dst


def find_pdf_page(doc, target):
    """找印刷页码 target 对应的 PDF 页（0-based）。

    FR 每页的印刷页码独立成词，位于页眉（顶部）或页脚（底部）。
    难点：正文里也可能出现同值的数字（表格、条款号）。
    因此按置信度分级：
      A. 全页仅 1 处该数字，且位于顶部 12% 或底部 10% 区域 → 最可信
      B. 全页仅 1 处该数字（位置不限）
      C. 块级文本恰好等于该数字
    """
    tgt = str(target)
    cand_a = []
    for i, pg in enumerate(doc):
        ws = [w for w in pg.get_text("words") if w[4].strip().strip(".,") == tgt]
        if len(ws) != 1:
            continue
        h = pg.rect.height
        w0 = ws[0]
        if w0[1] < 0.12 * h or w0[3] > 0.90 * h:
            cand_a.append(i)
    if len(cand_a) == 1:
        return cand_a[0]
    if cand_a:                       # 多页命中时取最靠前的一页
        return cand_a[0]

    for i, pg in enumerate(doc):
        ws = [w for w in pg.get_text("words") if w[4].strip().strip(".,") == tgt]
        if len(ws) == 1:
            return i

    for i, pg in enumerate(doc):
        for b in pg.get_text("blocks"):
            if b[4].strip() == tgt:
                return i
    return None


def main(only=None):
    import pymupdf as fitz
    sys.path.insert(0, str(ROOT))
    from fr_column_extract import extract_page_columns

    CORR.mkdir(parents=True, exist_ok=True)
    done, failed = [], []
    for vol, page, date, secs, mb in TARGETS:
        if only and not (str(vol) == only[0] and str(page) == only[1]):
            continue
        cite = f"{vol} FR {page}"
        print(f"\n=== {cite}  ({date})  影响 {secs}")
        try:
            pdf = download(date, mb)
        except Exception as e:
            print(f"    !! 下载失败 {e}")
            failed.append((cite, f"下载失败 {e}"))
            continue
        try:
            doc = fitz.open(pdf)
        except Exception as e:
            print(f"    !! 打开失败 {e}")
            failed.append((cite, f"打开失败 {e}"))
            continue

        pi = find_pdf_page(doc, page)
        if pi is None:
            print(f"    !! 未定位到印刷页 {page}")
            failed.append((cite, "未定位到页码"))
            doc.close()
            continue

        # 取该页及次页（更正件可能跨页）
        lines = []
        for j in (pi, pi + 1):
            if j < doc.page_count:
                lines.append(f"\n[[PDF page {j+1}]]")
                lines.extend(extract_page_columns(doc[j]))
        total = doc.page_count
        doc.close()

        txt = "\n".join(lines)
        head = (f"[[Federal Register Vol. {vol}, p.{page} — {date}]]\n"
                f"(来源：govinfo FR-{date} 整期影印 PDF 第 {pi+1} 页，OCR 文字层 + 分栏重排)\n"
                f"(整期共 {total} 页，本更正件还在次页时一并纳入)\n\n")
        stem = f"Correction_{vol}FR{page}_{date}"
        out = CORR / f"{stem}.clean.txt"
        out.write_text(head + txt.strip() + "\n", encoding="utf-8")
        (CORR / f"{stem}_meta.json").write_text(json.dumps({
            "citation": cite, "publication_date": date,
            "source": f"govinfo FR-{date} 整期影印 PDF p.{pi+1}",
            "affected_sections": secs, "extraction": "pymupdf + 分栏重排",
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"    -> {out.name}  ({len(txt):,} chars)")
        done.append(cite)

    print("\n" + "=" * 70)
    print(f"完成 {len(done)} 份: {', '.join(done)}")
    if failed:
        print("失败:")
        for c, w in failed:
            print(f"    {c}: {w}")


if __name__ == "__main__":
    args = sys.argv[1:]
    main((args[0], args[1]) if len(args) >= 2 else None)
