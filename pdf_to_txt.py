#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T3 / T4 — 给 Rulemaking Docs/ 里没有文字层的文献补 .txt

两类目标：
T3 本地更正件 PDF（58 FR 51970 / 43 FR 52495）+ govinfo 整期影印里的 30 FR 258
    → 落到 Rulemaking Docs/Corrections/
  T4  扫描全库，找出「有 PDF 或 HTML、但没有 .txt」的文件，逐个抽文字层

抽取优先级：PyMuPDF(fitz) > pypdf > html 去标签
用法：
    python pdf_to_txt.py --scan      # 只扫描，列出缺口
    python pdf_to_txt.py --apply     # 实际抽取落盘
"""
import html
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "Rulemaking Docs"
CORR = DOCS / "Corrections"

# T3 指定的两份更正件（相对 Rulemaking Docs）
CORRECTION_PDFS = [
    ("FinalRule_Docket 26269_Amdt23-45_58FR42136_1993-08-06_0002.pdf",
     "Correction_58FR51970_1993-10-05_Amdt23-45"),
    ("FinalRule_Docket 14324_ 14606_ 14625_ 14685_ 14779_Amdt23-23_25-46_27-16_29_43FR50578_1978-10-30_0002.pdf",
     "Correction_43FR52495_1978-11-13_Amdt23-23"),
]

# 1994 年前 FR 只有 govinfo 整期影印 PDF，需要先下载整期再按页抽取。
# (期号 ISO, 该期 PDF 的临时存放路径, 目标 FR 引证, 1-based 页码列表, 输出 stem)
WHOLE_ISSUE_EXTRACTS = [
    ("FR-1965-01-09", "/tmp/FR-1965-01-09.pdf", "30 FR 258", [12],
     "Correction_30FR258_1965-01-09_Amdt23-0"),
]


# ---------------------------------------------------------------- 抽取器
def _backend():
    try:
        import fitz  # PyMuPDF
        return ("pymupdf", fitz)
    except Exception:
        pass
    try:
        import pypdf
        return ("pypdf", pypdf)
    except Exception:
        pass
    return (None, None)


def extract_pdf_text(pdf: Path) -> str:
    """PDF 文字层抽取。

    FR 影印件是三栏版面，pymupdf 默认的 get_text() 会按 PDF 绘制顺序把左右栏
    交错输出。这里统一走 fr_column_extract 的「竖切分栏 + 逐栏自上而下」重排；
    只有在它失败（单栏文件或非 wide 版面）时才退回原始顺序。
    """
    try:
        from fr_column_extract import extract_pdf as col_extract
        return col_extract(pdf)
    except Exception:
        pass
    name, mod = _backend()
    if name == "pymupdf":
        doc = mod.open(pdf)
        parts = [pg.get_text() for pg in doc]
        doc.close()
        return "\n".join(parts)
    if name == "pypdf":
        r = mod.PdfReader(str(pdf))
        return "\n".join((pg.extract_text() or "") for pg in r.pages)
    raise RuntimeError("既无 pymupdf 也无 pypdf")


def extract_html_text(p: Path) -> str:
    t = p.read_text(encoding="utf-8", errors="replace")
    t = re.sub(r"<script\b.*?</script>", "", t, flags=re.S | re.I)
    t = re.sub(r"<style\b.*?</style>", "", t, flags=re.S | re.I)
    t = re.sub(r"<head\b.*?</head>", "", t, flags=re.S | re.I)
    t = re.sub(r"<br\s*/?>", "\n", t, flags=re.I)
    t = re.sub(r"</(p|div|tr|h[1-6]|li)>", "\n", t, flags=re.I)
    t = re.sub(r"<[^>]+>", "", t)
    t = html.unescape(t)
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", t)
    t = "\n".join(ln.rstrip() for ln in t.splitlines())
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip() + "\n"


def normalize(t: str) -> str:
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", t)
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    t = "\n".join(ln.rstrip() for ln in t.splitlines())
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip() + "\n"


def meaningful(t: str, floor: int = 200) -> bool:
    """判断抽取结果是否真的有内容（防止扫描件抽空壳塞进库里）"""
    alpha = sum(c.isalpha() for c in t)
    return alpha >= floor


# ---------------------------------------------------------------- 扫描
def base_of(p: Path) -> str:
    """去掉 _0002 / _0003 之类的分卷后缀，得到文献主干"""
    s = p.stem
    if s.endswith(".clean"):
        s = s[: -len(".clean")]
    return re.sub(r"_\d{3,4}$", "", s)


def scan_gaps():
    buckets = {}
    for p in sorted(DOCS.rglob("*")):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext not in (".pdf", ".html", ".txt"):
            continue
        if p.parent == CORR:
            continue
        buckets.setdefault(base_of(p), {}).setdefault(ext, []).append(p)

    gaps = []
    for b, exts in sorted(buckets.items()):
        if ".txt" not in exts:
            srcs = exts.get(".pdf", []) + exts.get(".html", [])
            gaps.append((b, sorted(srcs, key=lambda x: (x.suffix != ".html", x.name))))
    return gaps


# ---------------------------------------------------------------- 主流程
def do_corrections():
    CORR.mkdir(parents=True, exist_ok=True)
    name, _ = _backend()
    print(f"[T3] PDF 抽取后端：{name}\n")
    for rel, stem in CORRECTION_PDFS:
        src = DOCS / rel
        print(f"--- {stem}")
        if not src.exists():
            print(f"    !! 源文件不存在：{rel}")
            continue
        try:
            t = normalize(extract_pdf_text(src))
        except Exception as e:
            print(f"    !! 抽取失败 {e}")
            continue
        print(f"    抽出 {len(t):,} chars，字母数 {sum(c.isalpha() for c in t):,}")
        if meaningful(t):
            out = CORR / f"{stem}.clean.txt"
            out.write_text(t, encoding="utf-8")
            print(f"    -> {out.name}")
        else:
            print("    !! 疑似扫描件/无有效文字层，未落盘")
    do_whole_issue()
    print()


def do_whole_issue():
    """从 govinfo 整期影印 PDF 按页抽取 1994 年前的更正件"""
    if not WHOLE_ISSUE_EXTRACTS:
        return
    try:
        import pymupdf as fitz
    except ImportError:
        print("[T3] 无 pymupdf，跳过整期抽取")
        return
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    try:
        from fr_column_extract import extract_page_columns
    except Exception as e:
        print(f"[T3] 无法载入分栏抽取器 {e}")
        return

    for issue, pdf_path, citation, pages, stem in WHOLE_ISSUE_EXTRACTS:
        print(f"--- {stem}  ({citation}, 整期 {issue})")
        src = Path(pdf_path)
        if not src.exists():
            print(f"    !! 整期 PDF 不在 {pdf_path}")
            print(f"       可重新下载：https://www.govinfo.gov/content/pkg/{issue}/pdf/{issue}.pdf")
            continue
        doc = fitz.open(src)
        lines = []
        for pno in pages:
            if 1 <= pno <= doc.page_count:
                lines.extend(extract_page_columns(doc[pno - 1]))
            else:
                print(f"    !! 页码 {pno} 越界（共 {doc.page_count} 页）")
        doc.close()
        t = normalize("\n".join(lines))
        print(f"    抽出 {len(t):,} chars")
        if not meaningful(t):
            print("    !! 无有效文字层，未落盘")
            continue
        out = CORR / f"{stem}.clean.txt"
        out.write_text(t, encoding="utf-8")
        print(f"    -> {out.name}")


def do_apply():
    name, _ = _backend()
    print(f"[T4] PDF 抽取后端：{name}\n")
    gaps = scan_gaps()
    print(f"共发现 {len(gaps)} 个主干缺 .txt\n")
    ok, empty, fail = [], [], []
    for b, srcs in gaps:
        got = None
        for src in srcs:
            try:
                if src.suffix.lower() == ".pdf":
                    t = normalize(extract_pdf_text(src))
                else:
                    t = extract_html_text(src)
            except Exception as e:
                print(f"  !! {src.name}: {e}")
                continue
            if meaningful(t):
                got = (src, t)
                break
            print(f"  ~~ {src.name}: 抽出 {len(t):,} chars 但无实质内容，换下一个源")
        if got is None:
            fail.append((b, [s.name for s in srcs]))
            continue
        src, t = got
        # 落盘位置：跟源文件同名同目录，扩展名改 .txt
        out = src.with_suffix(".txt")
        out.write_text(t, encoding="utf-8")
        ok.append((b, src.name, len(t)))
        print(f"  OK  {b}\n        {src.name} -> {out.name}  ({len(t):,} chars)")
    print("\n" + "=" * 74)
    print(f"成功 {len(ok)} / 空壳 {len(empty)} / 失败 {len(fail)}")
    if fail:
        print("\n未能抽出（需人工或 OCR）：")
        for b, names in fail:
            print(f"  - {b}")
            for n in names:
                print(f"        {n}")


def do_scan():
    gaps = scan_gaps()
    print(f"缺 .txt 的主干：{len(gaps)} 个\n")
    for b, srcs in gaps:
        print(f"  {b}")
        for s in srcs:
            print(f"        {s.name}")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "--scan"
    if mode == "--scan":
        do_scan()
    elif mode == "--apply":
        do_apply()
    elif mode == "--corrections":
        do_corrections()
    elif mode == "--all":
        do_corrections()
        do_apply()
    else:
        print(__doc__)
