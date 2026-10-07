#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""枚举 DRS 里所有引用 Part 23 的 Final Rules / NPRM，输出权威全集清单。"""
import json
import os
import sys
import time

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_rulemaking import SESSION, BASE, OUTDIR, META  # noqa: E402

DOCTYPES = [("CFRFRSFAR", 11, "FinalRule"), ("NPRM", 10, "NPRM")]
FILTER_FIELD = "CFR Part Reference"


def browse(dt_label, dt_id, page):
    # saveResult=true 可突破接口默认「最多 50 条」的限制，一次返回筛选后的全集
    body = {"docTypeId": dt_id, "docTypeUrlLabel": dt_label,
            "filters": {FILTER_FIELD: ["Part 23"]}, "page": page,
            "saveResult": True}
    for _ in range(4):
        try:
            r = SESSION.post(f"{BASE}/api/browse/doctype/{dt_label}/documents/metadatas",
                             json=body, timeout=90)
            if r.status_code == 200:
                return r.json()
        except requests.RequestException:
            time.sleep(3)
    return None


def main():
    all_rows = []
    for label, tid, kind in DOCTYPES:
        seen = set()
        page = 1
        total = None
        while True:
            d = browse(label, tid, page)
            if not d:
                print(f"[{label}] page {page} 失败，停止", flush=True)
                break
            total = d.get("documentListTotalCount")
            lst = d.get("documentList") or []
            if not lst:
                break
            for x in lst:
                uid = x.get("docUniqueId") or ""
                unid = uid.split(".")[0].upper()
                if not unid or unid in seen:
                    continue
                seen.add(unid)
                st = {s["metadataName"]: s["metadataValue"] for s in (x.get("subText") or [])}
                all_rows.append({
                    "kind": kind,
                    "unid": unid,
                    "docUniqueId": uid,
                    "drsId": x.get("id"),
                    "mimeType": x.get("mimeType"),
                    "multiFile": x.get("multiFileFlag"),
                    "status": x.get("status"),
                    "docket": st.get("Docket Number") or "",
                    "notice": st.get("Notice Number") or "",
                    "amendment": st.get("Amendment") or "",
                    "issueDate": st.get("Issue Date") or "",
                    "effectiveDate": st.get("Effective Date") or "",
                    "citation": st.get("Citation") or "",
                    "subject": (st.get("Subject Heading") or st.get("Subject") or ""),
                })
            print(f"[{label}] page {page}: 累计 {len(seen)}/{total}", flush=True)
            if len(seen) >= (total or 0):
                break
            page += 1
            time.sleep(0.6)
            if page > 60:
                break

    out = "/Users/glennchou/FAR 23/FAR23_Part23规则文件全集.csv"
    cols = ["kind", "unid", "docUniqueId", "drsId", "mimeType", "multiFile",
            "status", "docket", "notice", "amendment", "issueDate",
            "effectiveDate", "citation", "subject"]
    import csv
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(all_rows)
    print(f"\n共 {len(all_rows)} 条 -> {out}")
    for kind in ("FinalRule", "NPRM"):
        print(" ", kind, sum(1 for r in all_rows if r["kind"] == kind))


if __name__ == "__main__":
    main()
