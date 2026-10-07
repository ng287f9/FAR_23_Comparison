#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
联邦公报影印 PDF 的**分栏重排**文字抽取。

背景：govinfo / GPO 的三栏正文版面（Federal Register）用 page.get_text() 直接抽，
会按 PDF 内部绘制顺序把左右栏交错输出，读起来是乱的。pymupdf 的 sort=True 也修不好
跨栏情况，必须自己做「竖切分栏 + 逐栏自上而下」的重排。

用法：
    python fr_column_extract.py <pdf路径> [> out.txt]
"""
import sys
from pathlib import Path

import pymupdf as fitz


def _column_bounds(words, page_width, min_gap_ratio=0.03, valley_ratio=0.30):
    """返回每栏的 (x0, x1)。

    联邦公报是典型的三栏版面，但字距抖动会让 x 投影出现「低覆盖谷」而不是
    完全空白，所以不能只找 cover==0 的区间。两步走：
      1. 找 cover==0 且足够宽的间隙（干净版面）
      2. 找不到就退而求其次，按投影直方图找「谷底」——低于峰值 valley_ratio
         的连续区间，取最深的若干个作为切口
    """
    if not words:
        return []
    step = 2.0
    n = int(page_width / step) + 2
    cover = [0] * n
    for w in words:
        a = max(0, int(w[0] / step))
        b = min(n - 1, int(w[2] / step))
        for i in range(a, b + 1):
            cover[i] += 1
    x0s = min(w[0] for w in words)
    x1s = max(w[2] for w in words)

    peak = max(cover) or 1
    dev = page_width * 0.10  # 允许在 WS 起点附近浮动

    def _runs(pred):
        runs, i = [], 0
        while i < n:
            if pred(cover[i]):
                j = i
                while j < n and pred(cover[j]):
                    j += 1
                runs.append((i, j))
                i = j
            else:
                i += 1
        return runs

    # ---- 方案 1：真·空白列间隙（排除页边距——页边距必然是 cover==0 的大块，
    #      若不剔除，方案 1 永远"成功"，把整页当成一栏）
    def _interior(g):
        return g[0] > x0s and g[1] < x1s

    gaps = [(a * step, b * step) for a, b in _runs(lambda c: c == 0)
            if (b - a) * step >= max(page_width * min_gap_ratio, 8) and _interior((a * step, b * step))]
    # ---- 方案 2：投影谷底（column gutter）
    #     注意：谷底通常只有 6~15pt 宽，要求 >= page_width*0.03 会把它们全滤掉
    if not gaps:
        thr = peak * valley_ratio
        gaps = [(a * step, b * step) for a, b in _runs(lambda c: c <= thr)
                if (b - a) * step >= 6.0]

    cuts = [0.5 * (a + b) for a, b in gaps]
    bounds, prev = [], x0s - 2.0
    for c in cuts:
        if c - prev > dev:
            bounds.append((prev, c))
            prev = c
    bounds.append((prev, x1s + 2.0))
    return [(a, b) for a, b in bounds if b - a > dev]


def extract_page_columns(page, line_tol=3.0):
    """返回该页按「先栏、后行」顺序排好的文本行列表。"""
    words = page.get_text("words")
    if not words:
        return []
    pw = page.rect.width
    bounds = _column_bounds(words, pw)
    lines = []
    for c0, c1 in bounds:
        col = [w for w in words if c0 - 2 <= w[0] < c1]
        if len(col) < 2:
            lines.extend(x[4] for x in col)
            continue
        # 用「包围盒纵向重叠」分行，比直接比 baseline 稳：同一行里夹着
        # 上/下标、不同字号时 baseline 会差好几 pt，很容易错行。
        col.sort(key=lambda w: (w[1] + w[3]) / 2)
        groups = []
        for w in col:
            top, bot = w[1], w[3]
            placed = False
            for g in groups:
                if min(bot, g["bot"]) - max(top, g["top"]) > 0.45 * min(
                    bot - top, g["bot"] - g["top"]
                ):
                    g["ws"].append(w)
                    g["top"] = min(g["top"], top)
                    g["bot"] = max(g["bot"], bot)
                    placed = True
                    break
            if not placed:
                groups.append({"ws": [w], "top": top, "bot": bot})
        for g in sorted(groups, key=lambda g: g["top"]):
            g["ws"].sort(key=lambda w: w[0])
            lines.append(" ".join(x[4] for x in g["ws"]))
    return lines


def extract_pdf(pdf: Path) -> str:
    doc = fitz.open(pdf)
    parts = []
    for pi, page in enumerate(doc):
        lines = extract_page_columns(page)
        parts.append(f"\n[[Page {pi + 1}]]\n" + "\n".join(lines))
    doc.close()
    return "\n".join(parts) + "\n"


if __name__ == "__main__":
    p = Path(sys.argv[1])
    sys.stdout.write(extract_pdf(p))
