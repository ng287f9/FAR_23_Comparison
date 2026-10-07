#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T2b — 清洗 Corrections/ 里 federalregister raw_text 抓回的 HTML 包装文本

federalregister.gov 的 full_text/text/*.txt 实际返回 HTML（<pre> 包正文），
这里剥掉标签、还原实体、去掉 Cloudflare 邮件混淆脚本，落 .clean.txt
"""
import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CDIR = ROOT / "Rulemaking Docs" / "Corrections"


def clean(raw: str) -> str:
    t = raw
    # 去 script
    t = re.sub(r"<script\b.*?</script>", "", t, flags=re.S | re.I)
    t = re.sub(r"<head\b.*?</head>", "", t, flags=re.S | re.I)
    # Cloudflare email protection 占位 → 统一文案
    t = re.sub(r'<a href="/cdn-cgi/l/email-protection".*?</a>', "[email protected]", t, flags=re.S | re.I)
    # 去所有标签
    t = re.sub(r"<[^>]+>", "", t)
    t = html.unescape(t)
    # 去 NUL 及其他控制字符（federalregister 文本层偶发 \x00）
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", t)
    # 行尾空白
    t = "\n".join(ln.rstrip() for ln in t.splitlines())
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip() + "\n"


def main():
    for p in sorted(CDIR.glob("*.txt")):
        if p.stem.endswith(".clean"):
            continue
        raw = p.read_text(encoding="utf-8", errors="replace")
        out = clean(raw)
        dst = p.with_suffix("").with_suffix(".clean.txt")
        # with_suffix("") 去掉 .txt；再补 .clean.txt
        dst = p.parent / (p.stem + ".clean.txt")
        dst.write_text(out, encoding="utf-8")
        print(f"{p.name:<52} {len(raw):>7,} -> {len(out):>7,} chars")


if __name__ == "__main__":
    main()
