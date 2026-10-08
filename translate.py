#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文化翻译层 —— 把 NPRM / Final Rule 的英文讨论段落译成中文。

后端：本机 FreeLLMAPI 的 OpenAI 兼容路由（127.0.0.1:31415），统一 key 从
`~/Library/Application Support/FreeLLMAPI/freeapi.db` 的 settings 表读，不落盘明文。

特性：
  · JSON 缓存（按「原文 md5 + 语种方向」索引），重复跑零成本
  · 术语表注入，保证 Part 23 术语前后一致
  · 长文本按句切块，块间并发，失败重试 + 退避
  · 译文只保留中文，不夹带解释、不夹带英文原句

用法：
    python translate.py --stats                 # 看缓存量
    python translate.py --test                  # 单条冒烟测试
    python translate.py --file in.txt           # 翻译整个文件（按空行分段）
"""
import argparse
import concurrent.futures as cf
import hashlib
import json
import os
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.request

ROOT = os.path.dirname(os.path.abspath(__file__))
CACHE_PATH = os.path.join(ROOT, "FAR23_译文缓存.json")
DB = os.path.expanduser("~/Library/Application Support/FreeLLMAPI/freeapi.db")
# 端点清单（含密钥，已 gitignore）。由 setup_endpoints.py 生成，也可手工编辑。
ENDPOINTS_FILE = os.path.join(ROOT, "translate_endpoints.json")

# 端点与密钥可通过环境变量覆盖，便于临时换成任意 OpenAI 兼容服务：
#   export TRANSLATE_BASE="https://api.deepseek.com/v1/chat/completions"
#   export TRANSLATE_KEY="sk-xxxx"
#   export TRANSLATE_MODEL="deepseek-chat"
BASE = os.environ.get("TRANSLATE_BASE", "http://127.0.0.1:31415/v1/chat/completions")

# ------------------------------------------------------------------ 术语表
# 注入到提示词里，保证同一术语在不同条款、不同批次中译法一致
GLOSSARY = [
    ("airworthiness standards", "适航标准"),
    ("airplane", "飞机"),
    ("commuter category", "通勤类"),
    ("normal category", "正常类"),
    ("utility category", "实用类"),
    ("acrobatic category", "特技类"),
    ("certification basis", "取证基础"),
    ("type certificate", "型号合格证"),
    ("airworthiness directive", "适航指令"),
    ("special condition", "专用条件"),
    ("notice of proposed rulemaking", "规章制定提案通告"),
    ("final rule", "最终规章"),
    ("preamble", "前言"),
    ("supplementary information", "补充说明"),
    ("amendatory instruction", "修订指令"),
    ("docket", "案卷"),
    ("commenter", "评论者"),
    ("petition", "请愿书"),
    ("exemption", "豁免"),
    ("harmonization", "协调统一"),
    ("JAA", "欧洲联合航空局（JAA）"),
    ("JAR", "欧洲联合航空要求（JAR）"),
    ("ARAC", "航空规则制定咨询委员会（ARAC）"),
    ("ARC", "航空规则制定委员会（ARC）"),
    ("GAMA", "通用航空制造商协会（GAMA）"),
    ("NTSB", "美国国家运输安全委员会（NTSB）"),
    ("NASA", "美国国家航空航天局（NASA）"),
    ("OFR", "联邦公报局（OFR）"),
    ("FAA", "美国联邦航空局（FAA）"),
    ("Federal Register", "《联邦公报》"),
    ("Code of Federal Regulations", "《联邦规章汇编》"),
    ("part 23", "第 23 部"),
    ("paragraph", "段"),
    ("subparagraph", "子段"),
    ("load factor", "载荷因数"),
    ("limit load", "限制载荷"),
    ("ultimate load", "极限载荷"),
    ("factor of safety", "安全系数"),
    ("damage tolerance", "损伤容限"),
    ("fatigue evaluation", "疲劳评定"),
    ("safe life", "安全寿命"),
    ("fail-safe", "破损安全"),
    ("gust", "阵风"),
    ("maneuvering", "机动"),
    ("emergency landing", "应急着陆"),
    ("dynamic conditions", "动力情况"),
    ("inertia force", "惯性力"),
    ("occupant", "乘员"),
    ("restraint system", "约束系统"),
    ("safety belt", "安全带"),
    ("harness", "束缚装置"),
    ("turnover", "翻转"),
    ("mass item", "质量块"),
    ("static test", "静力试验"),
    ("dynamic test", "动力试验"),
    ("flutter", "颤振"),
    ("pressurized cabin", "增压座舱"),
    ("cabin", "座舱"),
    ("cargo compartment", "货舱"),
    ("firewall", "防火墙"),
    ("fire extinguishing", "灭火"),
    ("flammable fluid", "可燃液体"),
    ("fuel tank", "燃油箱"),
    ("Instructions for Continued Airworthiness", "持续适航文件"),
    ("placard", "标牌"),
    ("flight manual", "飞行手册"),
    ("operating limitations", "使用限制"),
    ("kinds of operation", "运行种类"),
    ("maximum takeoff weight", "最大起飞重量"),
    ("stalling speed", "失速速度"),
    ("minimum control speed", "最小操纵速度"),
    ("takeoff", "起飞"),
    ("landing", "着陆"),
    ("balked landing", "中断着陆"),
    ("accelerate-stop", "加速—停止"),
    ("propeller", "螺旋桨"),
    ("turbine engine", "涡轮发动机"),
    ("reciprocating engine", "活塞式发动机"),
    ("auxiliary power unit", "辅助动力装置"),
    ("thrust reverser", "反推装置"),
    ("ice protection", "结冰防护"),
    ("lightning protection", "雷电防护"),
    ("HIRF", "高强度辐射场（HIRF）"),
    ("cockpit voice recorder", "驾驶舱话音记录器"),
    ("flight recorder", "飞行记录器"),
    ("emergency exit", "应急出口"),
    ("emergency evacuation", "应急撤离"),
    ("ditching", "水上迫降"),
    ("oxygen equipment", "氧气设备"),
    ("deicing", "除冰"),
    ("seaplane", "水上飞机"),
    ("float", "浮筒"),
    ("hull", "船身"),
    ("skiplane", "滑橇式飞机"),
    ("landing gear", "起落架"),
    ("nose wheel", "前轮"),
    ("tail wheel", "尾轮"),
    ("shock absorption", "减震"),
    ("drop test", "落震试验"),
    ("economic evaluation", "经济性评估"),
    ("regulatory flexibility", "监管灵活性"),
    ("Paperwork Reduction Act", "《文书削减法》"),
    ("unfunded mandate", "无经费强制要求"),
    ("effective date", "生效日期"),
    ("compliance", "符合性"),
    ("applicant", "申请人"),
    ("Administrator", "局长"),
]

SYSTEM = """你是民航适航规章（14 CFR Part 23）领域的资深中英翻译。
任务：把美国联邦航空局（FAA）规章制定文件（NPRM / Final Rule）中的英文段落译成简体中文。

硬性要求：
1. 只输出译文，不要任何前言、说明、注释、引号包裹或 Markdown 标记。
2. 忠实原文，不增不减不推测。原文没说的绝不补充。
3. 保留所有条款号、段号、编号的原文形式（如 Sec. 23.561(b)(2)(iv)、(a)、paragraph (b)），
   但把 Sec. / section 译作"条"，paragraph 译作"段"。例：Sec. 23.561(b) → §23.561(b) 段。
4. 保留所有数字、单位、阈值、日期、FR 引证（如 52 FR 1806）、案卷号（Docket No. 23516）原样。
5. 引号内的英文法规原文短句译为中文，但专业缩写（HIC、ATD、g、VSO 等）保留原文。
6. 语气客观、书面，用"美国联邦航空局（FAA）"指代 FAA；用"本局"指代 FAA 自称 we 时须谨慎，
   原文用 we 时译作"美国联邦航空局"。
7. 遇 "[Reserved]"、"[Deleted]" 保留方括号并译作"[预留]"、"[已删除]"。
8. 术语严格按下表统一：
{glossary}
"""

USER_TPL = """把下面的英文规章段落译成中文。只输出译文。

{text}"""

# 本机 FreeLLMAPI 那套模型名（走 127.0.0.1:31415）时的路由链，仅作兜底
MODEL_CHAIN = ["deepseek-v4-flash:free", "deepseek/deepseek-v4-flash-free",
               "glm-5.3", None]

# 单次请求的最大生成长度（切块默认 1100 字符，对应约 2400 token 足够）
MAX_TOKENS = int(os.environ.get("TRANSLATE_MAX_TOKENS", "2400"))

# 模型偶尔会先输出一段"思考/分析"再给译文，这些是典型开场白
LEAD_JUNK = re.compile(
    r"^(?:the user wants|let me|i need to|i will|here is|here's|ok[,.]|"
    r"sure[,.]|certainly|first,|分析任务|分析原文|思考过程|译文如下|翻译结果)", re.I)
CJK = re.compile(r"[\u4e00-\u9fff]")


def _clean(out):
    """剥掉模型夹带的前言/思考/标记，只留译文。"""
    if not out:
        return ""
    s = out.strip()
    s = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", s).strip()
    # 去掉包裹引号
    if len(s) > 2 and s[0] in "\"'“”" and s[-1] in "\"'“”":
        s = s[1:-1].strip()
    lines = [l.rstrip() for l in s.splitlines()]
    # 跳过开头的元叙述行（英文分析、编号列表、"译文："标记）
    i = 0
    while i < len(lines):
        l = lines[i].strip()
        if not l:
            i += 1
            continue
        if LEAD_JUNK.match(l) or re.match(r"^\d+[.、)]\s*(\*\*)?(分析|任务|原文|步骤|翻译)", l):
            i += 1
            continue
        if l.startswith("**") and ("分析" in l or "任务" in l or "原文" in l):
            i += 1
            continue
        if re.match(r"^[-#*]+\s*(分析|任务|原文|说明)", l):
            i += 1
            continue
        break
    lines = lines[i:]
    # 若正文里有明确的"译文："分隔，取其后
    for j, l in enumerate(lines):
        if re.match(r"^\s*(译文|翻译结果|最终结果)\s*[:：]\s*\S", l):
            lines = [re.sub(r"^\s*(译文|翻译结果|最终结果)\s*[:：]\s*", "", l)] + lines[j + 1:]
            break
    s = "\n".join(lines).strip()
    s = s.replace("**", "")
    return s


def zh_ratio(s):
    if not s:
        return 0.0
    return len(CJK.findall(s)) / max(len(s), 1)


def is_valid_zh(out, src=None):
    """判定是否像真正的中文译文（而不是一段英文分析）。"""
    s = (out or "").strip()
    if not s:
        return False
    if zh_ratio(s) < 0.25:
        return False
    if src and len(s) > len(src) * 2.2:
        return False
    return True


# ------------------------------------------------------------------ 后端
# 支持两类接口形态：
#   api="openai"  POST /v1/chat/completions  → choices[0].message.content
#   api="ollama"  POST /api/chat             → message.content（须 stream:false）
# 端点清单来自 translate_endpoints.json（gitignore），格式：
#   [{"name":..., "api":"ollama", "base":..., "key":..., "models":[...], "cap":8}, ...]
def _api_key():
    """环境变量 TRANSLATE_KEY → FreeLLMAPI 本地库。"""
    k = os.environ.get("TRANSLATE_KEY")
    if k:
        return k
    try:
        con = sqlite3.connect("file:" + DB + "?mode=ro", uri=True)
        v = con.execute("SELECT value FROM settings WHERE key='unified_api_key'").fetchone()
        con.close()
        return v[0] if v else None
    except Exception:
        return None


def _load_endpoints():
    """按优先级装配端点列表；返回 [] 表示只能用本机 FreeLLMAPI。"""
    # 1) 环境变量单点覆盖
    if os.environ.get("TRANSLATE_BASE") and os.environ.get("TRANSLATE_KEY"):
        return [{"name": "env", "api": "openai", "base": os.environ["TRANSLATE_BASE"],
                 "key": os.environ["TRANSLATE_KEY"],
                 "models": [os.environ.get("TRANSLATE_MODEL") or None],
                 "cap": int(os.environ.get("TRANSLATE_CAP", "8"))}]
    # 2) 端点清单文件
    if os.path.exists(ENDPOINTS_FILE):
        try:
            eps = json.load(open(ENDPOINTS_FILE, encoding="utf-8"))
            eps = [e for e in eps if e.get("base") and e.get("key")]
            if eps:
                return eps
        except Exception:
            pass
    # 3) 本机 FreeLLMAPI 路由
    k = _api_key()
    if k and os.environ.get("TRANSLATE_ALLOW_LOCAL", "0") != "0":
        return [{"name": "freellmapi", "api": "openai", "base": BASE, "key": k,
                 "models": list(MODEL_CHAIN), "cap": 8}]
    return []


def _post(ep, messages, max_tokens, timeout=150):
    """按端点形态发一次请求，返回模型输出文本。"""
    api = ep.get("api", "openai")
    if api == "ollama":
        payload = {"model": ep["model"], "stream": False, "messages": messages,
                   "options": {"temperature": 0.2, "num_predict": max_tokens}}
        if ep.get("extra"):
            payload.update(ep["extra"])
    else:
        payload = {"model": ep["model"], "messages": messages,
                   "temperature": 0.2, "max_tokens": max_tokens}
        if ep.get("extra"):
            payload.update(ep["extra"])
    req = urllib.request.Request(
        ep["base"], data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": "Bearer " + ep["key"],
                 "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        j = json.loads(r.read().decode("utf-8"))
    if api == "ollama":
        return (j.get("message") or {}).get("content") or ""
    ch = (j.get("choices") or [{}])[0]
    msg = ch.get("message") or {}
    c = msg.get("content")
    if isinstance(c, list):          # Cohere v2 等分段式 content
        c = "".join(x.get("text", "") for x in c)
    return c or ""


def _system_prompt():
    g = "\n".join(f"  · {en} = {zh}" for en, zh in GLOSSARY)
    return SYSTEM.format(glossary=g)


# ------------------------------------------------------------------ 切块
SENT_END = re.compile(r"(?<=[.!?])\s+(?=[A-Z\(“\"])|(?<=[。！？])")


def chunk(text, maxlen=1100):
    """按句切块；单句超长则硬切。"""
    text = text.strip()
    if len(text) <= maxlen:
        return [text]
    parts, cur = [], ""
    for s in SENT_END.split(text):
        s = s or ""
        if not s:
            continue
        if cur and len(cur) + len(s) + 1 > maxlen:
            parts.append(cur.strip())
            cur = s
        else:
            cur = (cur + " " + s).strip() if cur else s
    if cur:
        parts.append(cur.strip())
    out = []
    for p in parts:
        while len(p) > maxlen * 1.6:
            out.append(p[:maxlen])
            p = p[maxlen:]
        out.append(p)
    return [x for x in out if x.strip()]


# ------------------------------------------------------------------ 缓存
class Translator:
    def __init__(self, model=None, workers=8, verbose=False):
        import threading
        eps = _load_endpoints()
        if not eps:
            raise RuntimeError(
                "没有任何可用翻译端点。三选一：\n"
                "  1) 运行 python setup_endpoints.py（从本地密钥库生成 translate_endpoints.json）\n"
                "  2) export TRANSLATE_BASE=<OpenAI 兼容端点> TRANSLATE_KEY=<密钥> "
                "[TRANSLATE_MODEL=<模型名>]\n"
                "  3) 启动本机 FreeLLMAPI 路由（默认读其 unified_api_key）")
        if model:
            eps = [dict(e, models=[model]) for e in eps[:1]]
        # 摊平成 (端点, 模型) 路由槽，顺序即优先级
        self.slots = []
        for e in eps:
            for m in (e.get("models") or [None]):
                self.slots.append(dict(e, model=m))
        # 每端点并发闸门，避免把免费额度打爆
        self.sems = {}
        for e in eps:
            self.sems[e["name"]] = threading.Semaphore(int(e.get("cap", 8)))
        self.workers = workers
        self.verbose = verbose
        self.cache = {}
        if os.path.exists(CACHE_PATH):
            try:
                self.cache = json.load(open(CACHE_PATH, encoding="utf-8"))
            except Exception:
                self.cache = {}
        self._dirty = False
        self.calls = 0

    @staticmethod
    def _k(text):
        return hashlib.md5(text.encode("utf-8")).hexdigest()

    def save(self):
        if self._dirty:
            json.dump(self.cache, open(CACHE_PATH, "w", encoding="utf-8"),
                      ensure_ascii=False, indent=0)
            self._dirty = False

    # ---- 单块翻译（端点/模型路由链 + 退避重试 + 清洗 + 合法性校验）----
    def _one(self, text):
        messages = [
            {"role": "system", "content": _system_prompt()},
            {"role": "user", "content": USER_TPL.format(text=text)},
        ]
        last = None
        for ep in self.slots:
            sem = self.sems.get(ep["name"])
            for attempt in range(3):
                try:
                    if sem:
                        sem.acquire()
                    try:
                        raw = _post(ep, messages, MAX_TOKENS)
                    finally:
                        if sem:
                            sem.release()
                    self.calls += 1
                    out = _clean(raw)
                    if is_valid_zh(out, text):
                        return out
                    last = f"译文不合法（{ep['name']}/{ep['model'] or 'auto'}）：{out[:60]}"
                except urllib.error.HTTPError as e:
                    last = f"{ep['name']} HTTP {e.code}"
                    try:
                        last += " " + e.read().decode("utf-8", "replace")[:120]
                    except Exception:
                        pass
                except Exception as e:
                    last = f"{ep['name']} {e!r}"
                # 429/5xx 退避更久，普通错误快速重试
                wait = (2.5 if "HTTP 429" in str(last) or "HTTP 5" in str(last)
                        else 1.0) * (attempt + 1)
                time.sleep(wait)
        if self.verbose:
            print(f"    !! 翻译失败（{last}）：{text[:60]}…", file=sys.stderr)
        return None

    def purge_bad(self):
        """清掉缓存里不是中文译文的坏条目。"""
        bad = [k for k, v in self.cache.items() if not is_valid_zh(v)]
        for k in bad:
            del self.cache[k]
        if bad:
            self._dirty = True
            self.save()
        return len(bad)

    # ---- 整段翻译（切块 + 并发 + 缓存）----
    def translate(self, text):
        text = (text or "").strip()
        if not text:
            return ""
        parts = chunk(text)
        todo, res = [], [""] * len(parts)
        for i, p in enumerate(parts):
            k = self._k(p)
            if k in self.cache:
                res[i] = self.cache[k]
            else:
                todo.append((i, p, k))
        if todo:
            with cf.ThreadPoolExecutor(max_workers=self.workers) as ex:
                futs = {ex.submit(self._one, p): (i, k) for i, p, k in todo}
                for fu in cf.as_completed(futs):
                    i, k = futs[fu]
                    try:
                        out = fu.result()
                    except Exception:
                        out = None
                    if out:
                        res[i] = out
                        self.cache[k] = out
                        self._dirty = True
                    else:
                        res[i] = parts[i]  # 失败则退回英文，保证不丢内容
        return "\n".join(x for x in res if x).strip()

    def translate_many(self, texts, save_every=40):
        """真正并发的批量翻译：把所有文本的切块摊平后一次性提交，
        避免逐块串行（实测串行只有 2 条/分钟，并发后可到 40+ 条/分钟）。"""
        texts = [(t or "").strip() for t in texts]
        slot = []           # (text_idx, chunk_idx, cache_key, chunk_text)
        res = [None] * len(texts)
        for ti, t in enumerate(texts):
            if not t:
                res[ti] = ""
                continue
            parts = chunk(t)
            res[ti] = [""] * len(parts)
            for ci, p in enumerate(parts):
                slot.append((ti, ci, self._k(p), p))

        todo = []
        for ti, ci, k, p in slot:
            if k in self.cache:
                res[ti][ci] = self.cache[k]
            else:
                todo.append((ti, ci, k, p))

        if todo:
            done = 0
            with cf.ThreadPoolExecutor(max_workers=self.workers) as ex:
                futs = {ex.submit(self._one, p): (ti, ci, k, p) for ti, ci, k, p in todo}
                for fu in cf.as_completed(futs):
                    ti, ci, k, src = futs[fu]
                    try:
                        out = fu.result()
                    except Exception:
                        out = None
                    if out:
                        res[ti][ci] = out
                        self.cache[k] = out
                        self._dirty = True
                    else:
                        res[ti][ci] = src      # 失败退回英文，保证不丢内容
                    done += 1
                    if save_every and done % save_every == 0:
                        self.save()
                        if self.verbose:
                            print(f"    …{done}/{len(todo)}（已调 {self.calls} 次）",
                                  file=sys.stderr)
            self.save()

        return ["\n".join(x for x in (r if isinstance(r, list) else [r]) if x).strip()
                if r is not None else "" for r in res]


# ------------------------------------------------------------------ CLI
SAMPLE = ("This proposal would amend Sec. 23.301(d) by limiting the applicability of "
          "Appendix A to \"single-engine, excluding turbines\" airplanes rather than the "
          "current single-engine limitation. The JAA proposed this change because turbine-"
          "powered airplanes have higher wing loadings and different gust response "
          "characteristics. No comments were received on this proposal, and it is adopted "
          "as proposed.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", action="store_true")
    ap.add_argument("--stats", action="store_true")
    ap.add_argument("--purge", action="store_true")
    ap.add_argument("--file")
    ap.add_argument("--model", default=None)
    ap.add_argument("--workers", type=int, default=8)
    a = ap.parse_args()

    t = Translator(model=a.model, workers=a.workers, verbose=True)

    if a.purge:
        print("清除坏条目:", t.purge_bad())
        return
    if a.stats:
        print(f"缓存条目 {len(t.cache)}，文件 {CACHE_PATH}")
        return
    if a.test:
        print("原文：\n", SAMPLE, "\n")
        print("译文：\n", t.translate(SAMPLE))
        t.save()
        return
    if a.file:
        raw = open(a.file, encoding="utf-8").read()
        paras = [p.strip() for p in re.split(r"\n\s*\n", raw) if p.strip()]
        for zh in t.translate_many(paras):
            print(zh, "\n")
        return
    ap.print_help()


if __name__ == "__main__":
    import argparse
    main()
