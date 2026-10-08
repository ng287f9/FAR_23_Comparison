#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 translate_endpoints.json —— 翻译后端端点清单。

从本机已存的平台密钥库（OneDrive 的 freellmapi-keys.json）里挑出实测可用的
端点，按优先级写成本项目的端点清单文件。该文件含明文密钥，已在 .gitignore 中排除。

用法：
    python setup_endpoints.py            # 生成/覆盖 translate_endpoints.json
    python setup_endpoints.py --probe    # 生成后逐端点冒烟测试
    python setup_endpoints.py --keys <path>   # 指定密钥库路径
"""
import argparse
import json
import os
import sys
import urllib.request

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "translate_endpoints.json")
KEYS = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-个人/90_工具软件/freellmapi-keys.json")

# 端点模板：api 形态 + 端点地址 + 候选模型（按实测速度/质量排序）+ 并发上限
TEMPLATE = [
    {"name": "ollama-cloud", "api": "ollama", "platform": "ollama",
     "base": "https://ollama.com/api/chat",
     "models": ["gemma4:31b", "nemotron-3-super"], "cap": 8},
    {"name": "openrouter", "api": "openai", "platform": "openrouter",
     "base": "https://openrouter.ai/api/v1/chat/completions",
     "models": ["google/gemma-4-31b-it:free",
                "nvidia/nemotron-3-super-120b-a12b:free"], "cap": 4},
    {"name": "zhipu", "api": "openai", "platform": "zhipu",
     "base": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
     "models": ["glm-4.5-flash"], "cap": 4,
     "extra": {"thinking": {"type": "disabled"}}},
    {"name": "mistral", "api": "openai", "platform": "mistral",
     "base": "https://api.mistral.ai/v1/chat/completions",
     "models": ["mistral-small-latest"], "cap": 4},
]

SAMPLE = ("This proposal would amend Sec. 23.301(d) by limiting the applicability "
          "of Appendix A to single-engine airplanes. No comments were received.")


def load_keys(path):
    if not os.path.exists(path):
        raise SystemExit(f"找不到密钥库：{path}")
    d = json.load(open(path, encoding="utf-8"))
    return {k["platform"]: k["key"] for k in d.get("keys", [])}


def call(ep, text=SAMPLE, timeout=90):
    import translate as T
    msgs = [{"role": "system", "content": T._system_prompt()},
            {"role": "user", "content": text}]
    return T._clean(T._post(ep, msgs, T.MAX_TOKENS, timeout=timeout))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keys", default=KEYS)
    ap.add_argument("--probe", action="store_true")
    ap.add_argument("--only", default="", help="只写指定端点，逗号分隔")
    a = ap.parse_args()

    keys = load_keys(a.keys)
    eps = []
    for t in TEMPLATE:
        k = keys.get(t["platform"])
        if not k:
            print(f"  跳过 {t['name']}（密钥库无 {t['platform']}）")
            continue
        if a.only and t["name"] not in a.only.split(","):
            continue
        e = {x: t[x] for x in t if x != "platform"}
        e["key"] = k
        eps.append(e)

    json.dump(eps, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"已写入 {OUT}：{len(eps)} 个端点")
    for e in eps:
        print(f"  · {e['name']:14} {e['api']:7} cap={e['cap']}  {', '.join(e['models'])}")

    if a.probe:
        print("\n冒烟测试：")
        for e in eps:
            for m in e["models"]:
                ep = dict(e, model=m)
                try:
                    out = call(ep)
                    ok = "OK " if out else "空 "
                    print(f"  · {e['name']}/{m}: {ok}{out[:70]}")
                except Exception as ex:
                    det = ""
                    if hasattr(ex, "read"):
                        try:
                            det = ex.read().decode("utf-8", "replace")[:90]
                        except Exception:
                            pass
                    print(f"  · {e['name']}/{m}: FAIL {str(ex)[:40]} {det}")


if __name__ == "__main__":
    main()
