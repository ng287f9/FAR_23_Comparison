#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 FAA DRS 抓取 FAR Part 23 修订史背后的 NPRM / Final Rule 全文。

链路（逆向自 DRS Angular 前端）：
  1) 条款弹窗  GET /api/browse/documents/summary/{docId}?docTypeId=8
     -> docContent(base64 HTML) 里有 <A CLASS="document-link" VALUE="<Domino UNID>">
  2) 规则文件  GET /api/browse/documents/summaryguid/{大写UNID}
     -> 元数据 + docContent(base64 HTML/PDF二进制) + multiFileList
  3) 二进制    GET /api/content/alf/{id}

注意：必须先建立 cookie jar（curl -c ck.txt https://drs.faa.gov/browse），否则 403。
"""
import base64
import csv
import html
import json
import os
import re
import sys
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor

import requests

BASE = "https://drs.faa.gov"
COOKIE = "/tmp/ck.txt"
OUTDIR = "/Users/glennchou/FAR 23/Rulemaking Docs"
META = os.path.join(OUTDIR, "_meta")
LIST_CSV = "/Users/glennchou/FAR 23/FAR23_规则文件获取清单.csv"


def load_cookies():
    """读 curl 的 Netscape cookie jar。"""
    jar = requests.cookies.RequestsCookieJar()
    if not os.path.exists(COOKIE):
        return jar
    for line in open(COOKIE, encoding="utf-8", errors="replace"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) < 7:
            continue
        domain, _flag, path, secure, _exp, name, value = parts[:7]
        jar.set(name, value, domain=domain.lstrip("."), path=path)
    return jar


SESSION = requests.Session()
SESSION.cookies = load_cookies()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Referer": BASE + "/browse",
    "Accept": "application/json, text/plain, */*",
})


def api_get(path, timeout=60, retries=3):
    url = BASE + path
    for i in range(retries):
        try:
            r = SESSION.get(url, timeout=timeout)
            if r.status_code == 200:
                return r
            if r.status_code in (403, 429):
                time.sleep(3 + 3 * i)
                continue
            return r
        except requests.RequestException:
            time.sleep(2 + 2 * i)
    return None


def html_to_text(h):
    """极简 HTML -> 纯文本，保留段落。"""
    h = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", h)
    h = re.sub(r"(?is)<!--.*?-->", " ", h)
    h = re.sub(r"(?i)<br\s*/?>", "\n", h)
    h = re.sub(r"(?i)</(p|div|tr|h[1-6]|li)>", "\n\n", h)
    h = re.sub(r"(?i)</t[dh]>", "\t", h)
    h = re.sub(r"<[^>]+>", "", h)
    h = html.unescape(h)
    h = re.sub(r"[ \t\xa0]+", " ", h)
    h = re.sub(r"\n{3,}", "\n\n", h)
    return h.strip()


def safe(s, n=60):
    s = re.sub(r"[^\w\-\. \(\)]+", "_", str(s or "")).strip()
    return s[:n].strip() or "unnamed"


def build_prefix(rec, meta):
    """生成可识别的文件名前缀。"""
    typ = "FinalRule" if "Final" in (rec["类型"] or "") else "NPRM"
    num = safe(rec["编号"], 40)
    date = rec["签发日期"] or ""
    docket = safe((meta or {}).get("Docket Number"), 20)
    amdt = safe((meta or {}).get("Amendment"), 20)
    bits = [typ, num]
    if docket and docket.lower() not in num.lower():
        bits.append("Docket" + docket)
    if amdt and amdt.lower() != "unnamed":
        bits.append("Amdt" + amdt)
    bits.append(date)
    return "_".join(b for b in bits if b)


def fetch_one(rec, idx_total=None):
    tag = rec["编号"] + " " + rec["签发日期"]
    unid = rec.get("unid") or ""
    if not unid:
        return {"编号": rec["编号"], "ok": False, "why": "无UNID"}

    mpath = os.path.join(META, unid + ".json")
    if os.path.exists(mpath):
        d = json.load(open(mpath, encoding="utf-8"))
    else:
        r = api_get("/api/browse/documents/summaryguid/" + unid)
        if r is None or r.status_code != 200:
            return {"编号": rec["编号"], "ok": False,
                    "why": "summaryguid HTTP %s" % (r.status_code if r else "ERR")}
        try:
            d = r.json()
        except Exception:
            return {"编号": rec["编号"], "ok": False, "why": "json解析失败"}
        json.dump(d, open(mpath, "w", encoding="utf-8"), ensure_ascii=False)

    meta = d.get("metadatas") or {}
    prefix = build_prefix(rec, meta)
    mime = (d.get("mimeType") or "").lower()
    written = []

    # 主文件：HTML 全文走 docContent
    dc = d.get("docContent")
    if dc:
        try:
            raw = base64.b64decode(dc)
        except Exception:
            raw = b""
        if raw:
            name = d.get("docName") or (unid + ".html")
            ext = ".html" if name.lower().endswith((".html", ".htm")) else ".html"
            if "pdf" in mime:
                ext = ".pdf"
                fp = os.path.join(OUTDIR, prefix + ext)
                open(fp, "wb").write(raw)
            else:
                fp = os.path.join(OUTDIR, prefix + ext)
                try:
                    txt = raw.decode("utf-8")
                except UnicodeDecodeError:
                    txt = raw.decode("latin-1", errors="replace")
                open(fp, "w", encoding="utf-8").write(txt)
                # 同时导一份纯文本，便于检索
                open(os.path.join(OUTDIR, prefix + ".txt"), "w",
                     encoding="utf-8").write(html_to_text(txt))
            written.append(os.path.basename(fp))
    elif d.get("id"):
        r = api_get("/api/content/alf/" + d["id"], timeout=180)
        if r is not None and r.status_code == 200 and r.content:
            name = d.get("docName") or ""
            ext = os.path.splitext(name)[1] or (".pdf" if "pdf" in mime else ".bin")
            fp = os.path.join(OUTDIR, prefix + ext)
            open(fp, "wb").write(r.content)
            written.append(os.path.basename(fp))

    # 附加文件（更正件、HTML版等）
    for f in (d.get("multiFileList") or []):
        fid, fname = f.get("fileId"), f.get("fileName") or ""
        if not fid:
            continue
        suffix = ""
        m = re.search(r"\.(\d{4})\.", fname)
        if m and m.group(1) != "0001":
            suffix = "_" + m.group(1)
        r = api_get("/api/content/alf/" + fid, timeout=180)
        if r is not None and r.status_code == 200 and r.content:
            ext = os.path.splitext(fname)[1] or ".bin"
            fp = os.path.join(OUTDIR, prefix + suffix + ext)
            open(fp, "wb").write(r.content)
            written.append(os.path.basename(fp))
            if ext.lower() in (".html", ".htm"):
                try:
                    txt = r.content.decode("utf-8")
                except UnicodeDecodeError:
                    txt = r.content.decode("latin-1", errors="replace")
                open(os.path.join(OUTDIR, prefix + suffix + ".txt"), "w",
                     encoding="utf-8").write(html_to_text(txt))

    return {"编号": rec["编号"], "ok": bool(written), "files": written,
            "mime": mime, "docName": d.get("docName"),
            "why": "" if written else "无内容",
            "citation": meta.get("Citation"), "subject": meta.get("Subject")
                        or meta.get("Subject Heading")}


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    workers = int(sys.argv[2]) if len(sys.argv) > 2 else 4
    os.makedirs(OUTDIR, exist_ok=True)
    os.makedirs(META, exist_ok=True)

    rows = list(csv.DictReader(open(LIST_CSV, encoding="utf-8-sig")))
    for r in rows:
        m = re.search(r"/([0-9a-fA-F]{32})!OpenDocument", r.get("rgl_url") or "")
        r["unid"] = m.group(1).upper() if m else ""
    # 先抓影响面大的
    rows.sort(key=lambda r: -int(r.get("影响条款版本数") or 0))
    if limit:
        rows = rows[:limit]

    results = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for i, res in enumerate(ex.map(fetch_one, rows), 1):
            results.append(res)
            flag = "OK " if res["ok"] else "FAIL"
            print("[%3d/%d] %s %-28s %s" % (
                i, len(rows), flag, res["编号"][:28],
                ", ".join(res.get("files", []))[:80] or res.get("why", "")), flush=True)

    ok = sum(1 for r in results if r["ok"])
    print("\n完成: %d/%d 成功" % (ok, len(results)))
    json.dump(results, open(os.path.join(OUTDIR, "_fetch_report.json"), "w"),
              ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
