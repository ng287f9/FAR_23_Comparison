# FAR Part 23 修订史工程 — 状态总览与任务清单

> 最后更新：2026-10-06 20:20
> 工作目录：`/Users/glennchou/FAR 23`

---

## 一、项目目标

把 FAA DRS 历史快照 **「Part 23 @ Amdt 23-63 as of 08/29/2017 and earlier」** 中
**14 CFR Part 23 全部条款**的修订历史抽取、综合，产出带**目录**、**深度背景/原因剖析**、
**彩色条文演变 diff**、**带交叉索引的参考文献**的 Word 文档。

最终交付物的组织方式（已与用户确认）：

- **按 Subpart 分卷**（A/B/C/D/E/F/G + 附录），每卷一个 .docx
- **外加一卷总索引**（全局条款表 + 修订案表 + 文献反向索引）
- **先做 Subpart C 试点**，跑通后再铺开其余分部

---

## 二、数据基座（已完成 ✅）

### 2.1 条款清单

`FAR23_Amdt23-63_条款清单.csv`（1217 行）

- **412 个唯一条款号**，**1203 条版本记录**，平均 2.92 版/条款
- 涵盖 **63 个修订案**（23-1 … 23-63，含 `23-45C` 与 `Initial`）

| 分部 / 附录 | 条款数 | 版本数 | 平均 |
|---|---:|---:|---:|
| Subpart A – General | 3 | 10 | 3.33 |
| **Subpart B – Flight** | 52 | 191 | 3.67 |
| **Subpart C – Structure（试点）** | **72** | **175** | 2.43 |
| Subpart D – Design and Construction | 82 | 212 | 2.59 |
| Subpart E – Powerplant | 87 | 277 | 3.18 |
| Subpart F – Equipment | 57 | 176 | 3.09 |
| Subpart G – Operating Limitations | 31 | 113 | 3.65 |
| Appendix A | 7 | 16 | 2.29 |
| Appendix B | 2 | 6 | 3.00 |
| Appendix C / D / E | 各 1 | 各 2 | 2.00 |
| Appendix F | 3 | 5 | 1.67 |
| Appendix G | 4 | 6 | 1.50 |
| Appendix H | 6 | 6 | 1.00 |
| Appendix I / J | 各 1 | 各 1 | 1.00 |
| **合计** | **412** | **1203** | **2.92** |

### 2.2 规则制定文献（NPRM / Final Rule）

`Rulemaking Docs/` —— **368 个文件**（.html 151 / .txt 151 / .pdf 63 / _meta 151 JSON）

- 按文件名主干去重 **165 份**（FinalRule 87 / NPRM 78）
- 覆盖 1961–2022；**FinalRule 62/62 修订案全覆盖**、NPRM 52 已下载
- 9 个修订案 DRS 标 `NPRM Actions: Not Applicable`（本就不存在 NPRM）
- **唯一真实缺口**：Amdt 23-40 的 NPRM **Notice 88-9** —— DRS 服务端把 NPRM 记录
  错误指向了 Final Rule 文档（UNID `40244E92…`），非漏抓
- **15 份主干缺 .txt 文本层**（只有 PDF/HTML），批量生成前需处理

### 2.3 对应/覆盖关系表

| 文件 | 行数 | 用途 |
|---|---:|---|
| `FAR23_条款-规则文件对应表.csv` | 2914 | (Section × Amendment) → (类型, 编号, 日期, FR引证) |
| `FAR23_NPRM与FinalRule索引.csv` | 121 | 121 个唯一规则文件 + 影响条款版本数 |
| `FAR23_Part23规则文件全集.csv` | 150 | DRS 浏览库权威全集（含 UNID） |
| `FAR23_已下载规则文件清单.csv` | 365 | 本地文件 ↔ 类型/编号/Amdt/FR卷页 |
| `FAR23_修订案覆盖报告.csv` | 61 | 62 个修订案 × FinalRule/更正件/NPRM 状态 |
| `FAR23_规则文件获取清单.csv` | 121 | 含 rgl UNID、govinfo 链接、下载状态 |
| `FAR23_条款溯源注.csv` | 381 | CFR 溯源注原文 + 更正件判定 |
| `FAR23_更正通告清单.csv` | 19 | 更正件清单（按 FR 引证合并） |

---

## 三、已完成的关键工作

### 3.1 §23.561 试验文档（模板已定稿）✅

- 产出：`FAR23_Sec23.561_修订历史与背景分析.docx`（65 KB，123 段，18 表，约 2.7 万字）
- 生成脚本：`build_23_561_doc.py`（python-docx **直接构建**，不走 md→html→docx，
  因为需要逐 run 着色）
- 结构：封面 → 编制说明 → TOC 域 + 手工目录 → §1 修订历史汇总 → §2 逐版背景深剖
  → §3 条文原文演变对比（彩色 diff + 要素矩阵）→ §4 总结 → §5 参考文献（16 条）
  → 附录（DRS 记录 ID）
- **着色约定（脚本常量）**：蓝 `#0070C0`=新增 / 红 `#C00000`=修订 /
  灰 `#808080`+删除线=删除 / 黑 `#1A1A1A`=继承未变
- 参考文献 [1]–[16] 全部被正文引用，交叉索引闭合

**对参考样例 `part23_section561_amendment_history.docx` 的三处事实纠错**（样例有错，勿采信）：

1. 23.561 真实修订史只有 **7 版**：23-0(1965-02-01)、23-7(1969-09-14)、
   23-34(1987-02-17)、23-36(1988-09-14)、23-46(1994-06-16)、23-48(1996-03-11)、
   23-62(2012-01-31)。样例里的 23-14 / 23-19 / 23-43 / 23-64 **均不是** 23.561 的修订。
2. 源条款是 **CAR 3.386**（NPRM 64-17 里编号 Sec. 23.709 "Protection"），不是 3.385。
3. NPRM 64-17 提案 Upward = **2.0g**，最终规章 29 FR 17955 采用 **3.0g**。
4. **Amdt 23-34 没有改 23.561 的条文**：52 FR 1806 共 43 条修订指令不含 23.561；
   DRS 列它只是因为新增通勤类条款（23.785(g)/23.807(d)/23.963(f)）开始交叉引用它。

### 3.2 批量生成数据层 ✅

`part23_source.py`（506 行）—— 可复用的取证层，已通过 50/50 抽样解析验证

```
compact_paragraphs(path)    # 通用换行读取 + 空白压缩（抗 CR 行尾）
class Doc                   # path/text()/enrich()/label()/fr_citation()
build_doc_index()           # 133+ 篇文献索引（FinalRule 64 / NPRM 69）
load_sections(subpart)      # 按分部载入条款
load_correspondence()       # (条款, 修订号) → 文献
resolve_docs(num, kind, doc_index, secnum)
sec_regex / find_hits / grab_amendment / grab_discussion
mine_doc(doc, secnum)       # 挖出「修订指令」+「讨论/评论」两类段落
initial_clause_text(secnum) # 初始版条文
summarize_change(amend_block)
```

### 3.3 更正通告（Correction）问题排查 ✅ **本轮最新成果**

> ⚠ **本节为早期（risingup 口径）结论，已被 3.4 取代。**
> 权威口径见 `FAR23_更正通告权威台账.csv`（34 份）与 `FAR23_受影响条款清单.csv`（81 个条款）。
> 早期判定的 19 份只是其中被 CFR 溯源注显式列出的一部分。

> 用户提出的现象：溯源注里 `Doc. No.` 后面跟两个 FR 引证时，**第二个是第一个的更正通告**。
> 例：`[Doc. No. 26269, 58 FR 42161, Aug. 6, 1993; 58 FR 51970, Oct. 5, 1993]`

**判定规则**：CFR 溯源注里**没有 `Amdt.` 前缀、且不在任何已知修订案页码附近的裸 FR 引证**
＝ 更正通告。

**结果（381 个 risingup 页面全量扫描，页面缓存在 `/tmp/risingup_pages`）**：

- **66 / 381 条款（17.3%）带有更正件引证** ← 修正了初版误判的 35 条
- 合并为 **19 份独立更正件**（另 1 条是 risingup 把 §23.1105 的日期误印成
  "Jan. 9, 1996"，实为 30 FR 258 / 1965-01-09，已并入）

**影响面 Top 5**：

| 更正件 FR 引证 | 日期 | 影响条款数 | 涉及条款 |
|---|---|---:|---|
| **30 FR 258** | 1965-01-09 | **30** | 23.29/23.31/23.301/23.331/23.351/23.605/23.613/23.723/23.781/23.783/23.831/23.963/23.973/23.975/23.1105/23.1145/23.1193/23.1307/23.1329/23.1353/23.1357/23.1361/23.1387/23.1511/23.1521/23.1541/23.1543/23.1547/23.1555/23.1567 |
| **58 FR 51970** | 1993-10-05 | **8** | 23.525/23.527/23.533/23.535/23.573/23.629/23.775 + 附录 I |
| **52 FR 34745** | 1987-09-14 | **6** | 23.3/23.1199/23.1323/23.1351 + 附录 F/附录 G |
| **58 FR 27060** | 1993-05-06 | **5** | 23.961/23.971/23.1091/23.1191/23.1305 |
| **56 FR 5455** | 1991-02-11 | **3** | 23.161/23.423/23.701 |

其余：52 FR 7262(2)、53 FR 34194(2)、32 FR 13505 + 32 FR 13714(§23.1325 连着两封)、
34 FR 17509、38 FR 32784、55 FR 46888、55 FR 47028、35 FR 1102、
71 FR 537、73 FR 19746、73 FR 35063、74 FR 32799、74 FR 32800 各 1。

**按分部**：A 1 / B 4 / **C 11** / D 12 / E 12 / F 14 / G 8

**本地原文可得性（关键结论）**：

- 19 份里**只有 2 份本地真有正文**
  - **58 FR 51970**（23-45 Correction）：
    `Rulemaking Docs/FinalRule_Docket 26269_Amdt23-45_58FR42136_1993-08-06_0002.pdf`
    （3 页，DRS multiFile 名 `23-45 (Correction).pdf`，**仅有 PDF、未抽 txt**）
    已抽文本核实内容：把 §23.527(c)、§23.533(b)(1) 里的 "appendix H" 改为 "appendix I" 等
  - **68 FR 75390**（23-54 Correction, 2003-12-31）：
    `FinalRule_Docket FAA-1998-4815_Amdt23-54_33-20_68FR75390_2003-12-31.html/.txt`
    （**正文齐全**）
- 另有 **43 FR 52495**（1978-11-13，23-23 Correction）：
  `FinalRule_Docket 14324_ 14606_ 14625_ 14685_ 14779_Amdt23-23_..._0002.pdf`（2 页）
  内容为把 §23.785 标题 "Seats and berths." 更正为 "Seats, berths, safety belts, and harnesses."
  ⚠️ **这份更正件在 risingup 溯源注里完全查不到** —— 见下方"两个数据源的坑"
- `FinalRule_Docket 18334_..._0002/_0003.pdf`（54 FR 52932）只涉及 Part 91，**与 Part 23 无关**
- 其余 **12 份 1994 年前的更正件**：govinfo 只有整期影印 PDF（100–250 MB），OCR 成本高，暂不可得
- **5 份 1994 年后的**（71 FR 537 / 73 FR 19746 / 73 FR 35063 / 74 FR 32799 / 74 FR 32800）
  **可单篇获取**，通路已验证 —— 见"下一步 T2"

**两个数据源的坑（重要）**：

1. **DRS 基本不记录更正件**。`FAR23_条款-规则文件对应表.csv` 里 19 个更正件 FR 页码
   只命中 1 次（71 FR 537 → §23.773）。更正信息**只能靠 CFR 溯源注**。
2. **risingup 的溯源注不完整**。§23.785 的溯源注只写了 Amdt 23-36 / 23-49，
   **完全没有 Amdt 23-23**，也没有 43 FR 52495 更正件。
   → 66 条是**下限**，真实数量更多。

**已修掉的两个判定 bug**（`crawl_risingup.py` 已就地修正）：

- 站点用 **en dash（– U+2013）**，且写法有 `Amdt. 23–7` / `Amdt. No. 23–59` / `Amdt 23–34`
  三种 → 正则改为 `Amdt\.?\s*(?:No\.?\s*)?\s*(\d+)\s*[‐-―\-−]\s*(\d+[A-Z]?)`。
  （修复前 §23.905 被误判：73 FR 63345 实为 "Amdt. No. 23–59"，是修订案不是更正件）
- 形如 `…; 58 FR 51970, Oct. 5, 1993, as amended by Amdt. 23–48, 61 FR 5147, …`
  一段里**既有裸 FR 又有修订案**，旧逻辑被 Amdt 标记"带偏"而漏判。
  → 现在**同时按 `;` 和 `as amended by` 切分**。（修复后 §23.573 从 0 条变为 2 条更正件）

---

## 四、待决 / 待办任务

### T1 — 定更正件在成文里的呈现方式 ⬜ 待用户拍板

建议方案（待确认）：

- 年表（修订沿革表）里**单独一行**，类型列标「**更正通告 (Correction)**」，
  与「最终规章」「NPRM」区分
- 参考文献里单列一条，来源标注 `CFR 溯源注 / risingup.com`
- 若本地有正文 → 写出"更正了什么"；若无正文 → 注明
  「原文未获取（1994 年前 FR 仅存整期影印件）」，并说明该条款现行文本已包含更正结果

### T1 — 定更正件在成文里的呈现方式 ✅ 已按建议方案落地

采用方案（已写进 `FAR23_Sec23.561_修订历史与背景分析.docx` 2.8 / 2.3.1 验证通过）：

- 年表外**单列一节**「更正类文献」，类型分为
  **FAA 更正通告 / CFR 编纂校正 / 技术性修订** 三类
- 参考文献单列条目，来源标注 `52 FR 34745` 等原始 FR 引证 + 本地文件名
- 有影响条文的 → 写出"更正了什么"；无正文的 → 注明「原文未获取」
- ⚠ 新增一类必须排除项：**19 份更正件里只有 17 份真的影响 Part 23**，
  其余只修正他部（Part 91/33/35）编号或条文，成文时**不应误引**

### T2 — 抓取 1994 年后更正件正文 ✅（超额完成）

通路**已验证可用**（federalregister.gov API，1994 年起有数据）：

```bash
curl --compressed "https://www.federalregister.gov/api/v1/documents.json?per_page=1000&order=oldest&conditions[publication_date][is]=2008-04-11&fields[]=document_number&fields[]=title&fields[]=start_page&fields[]=end_page&fields[]=type&fields[]=raw_text_url"
```

**改用更强的检索条件**：`conditions[cfr][title]=14&conditions[cfr][part]=23`
直接锁定 Part 23（272 条全覆盖），比按 agency 扫快两个数量级。

成果（`fetch_corrections.py`，落 `Rulemaking Docs/Corrections/`）：

- **共识别 34 份**更正/校正类文献（原本以为只有 5 份 + 19 份）
- **已获取原文 30 份**；剩 4 份为 1994 年前（用户已指示不再处理）
- **新发现此前完全遗漏的 3 类**：
  - `C`- 前缀更正件（82 FR 2193 = C1-2016-28714，更正 Amdt 23-63）
  - **CFR Correction** 4 条（OFR 年度版排印校正，不进溯源注）——
    其中 **72 FR 59939 / 72 FR 72915 两条都针对 §23.561**
  - 61 FR 252（1996）§23.965 燃油箱试验振动循环折算
- 1994 年前的用 `fetch_old_corrections.py` 从 govinfo **整期影印 PDF 按印刷页抽取**
  （`fr_column_extract.py` 做双栏重排），成功补回 9 份，含 30 FR 258 / 58 FR 51970 / 43 FR 52495

### T3 — 抽本地更正件 PDF 的文字层 ✅

`pdf_to_txt.py --corrections`，`pymupdf` + 分栏重排，抽 58 FR 51970 / 43 FR 52495 / 30 FR 258。

### T4 — 处理 15 份缺 .txt 的文献 ✅

`pdf_to_txt.py --apply` 已跑完，`--scan` 复检缺 .txt 主干 = **0 个**。

### T5 — 解决 §23.561 的 23-24 / 23-34 分歧 ✅ **已查实并定案**

在 52 FR 34745 原文里找到了决定性证据：

- 该件 **FR 抬头**印作 `[Docket No. 23516; Amdt. Nos. 21-59, **23-24**, 36-13, 91-197, 135-21]`
- 但同一份文件的 `SUPPLEMENTARY INFORMATION` 与 `Correction of Publication`
  **两处都写 23-34**，且开篇明说更正 52 FR 1806（1987-01-15）
- 1987-03-09 同案卷更正件 52 FR 7261 抬头**正确印作 23-34**，可交叉印证

**→ 定案：FR 印刷错误，CFR 编纂照抄错误抬头，DRS 的 Amdt 23-34 是对的。**

**附带发现（更重要）**：52 FR 34745 第 9 项新增修订指令 25-1，
把 §23.561(b)(2) 表格第一栏标题 `Normal and utility categories`
改为 `Normal, utility, and commuter categories`。
→ **§23.561 的条文确实被这份更正件改过**，此前"23-34 未改条文"的结论需加限定。

已写入 `FAR23_Sec23.561_修订历史与背景分析.docx` 新增 2.3.1 与 2.8 节。

### 3.4 更正件权威台账（三源交叉验证）✅ **本轮最终成果**

脚本：`scan_corrections_official.py` / `fetch_corrections.py` / `fetch_old_corrections.py` /
`build_corrections_ledger.py`

**三源**：

| 源 | 覆盖 | 说明 |
|---|---|---|
| govinfo CFR **2016** 版 XML | 全部条款 | 2017 版已是改写后的新 Part 23，必须用 2016 版 |
| federalregister.gov API（`cfr[title]=14&cfr[part]=23`） | 1994 年后 | 含 `C`-前缀更正件与 CFR Correction |
| govinfo 整期影印 PDF 按印刷页抽取 | 1994 年前 | `fr_column_extract.py` 双栏重排 |

**产出**：

| 文件 | 内容 |
|---|---|
| `FAR23_更正通告权威台账.csv` | 34 份更正/校正文献，含性质 / 是否影响 Part 23 / 受影响条款 / 本地文件状态 |
| `FAR23_受影响条款清单.csv` | **81 个条款/附录**受影响，含"原文已获取 n/m" |
| `FAR23_溯源注日期勘误.csv` | 781 条 FR 引证的年份校验结果 |
| `Rulemaking Docs/Corrections/` | 30 份正文 |

**可复用的硬规律**：

1. **FR 卷号 + 1935 = 出版年份**（1965 年起恒定）。据此抓出 2 处 CFR 官方印刷错误：
   §23.929 的 33 FR 31822 应为 1968（官方印 1973）；
   §23.1105 的 30 FR 258 应为 1965（官方印 1996）。
2. **FR 文档号 `C`- 前缀 = 更正件**（如 C1-2016-28714）。
3. **CFR 溯源注里的"Amdt. 编号"可能是 FR 抬头的印刷错误**（§23.561 的 23-24 案，见 T5）。
4. 溯源注**会漏记**：带 `Amdt.` 前缀的更正件被判定规则当成修订案（52 FR 34745 因此少记 9 个条款）。
   → 台账里用 `VERIFIED_SECS`（逐份读原文）覆盖推断值。

### T6 — 生成 Subpart C 试点卷 ✅

产出：`FAR23_SubpartC_修订史.docx`（**191 KB / 72 条款 / 1686 段 / 234 表**）
脚本：`build_subpart_doc.py` + 新抽出的排版公共模块 `docx_layout.py`

**两种文档的定位分工（重要）**：

| | 精写版（`build_23_561_doc.py`） | 取证版（`build_subpart_doc.py`） |
|---|---|---|
| 范围 | 单条重点条款 | 分部全部条款 |
| 内容 | 逐版叙事剖析 + 彩色条文 diff + 要素矩阵 | 年表 + 修订指令原文 + 讨论摘录 + 更正件提示 + 参考文献 + 初始条文 |
| 生成方式 | 人工考据后写死 | 全自动，可一键铺开其他分部 |

取数层覆盖（Subpart C 实测）：

- **68 / 72** 条款的全部历史版本都能定位到 NPRM 或 Final Rule 原文
- **59 / 72** 挖到修订指令原文；**50 / 72** 挖到讨论/评论处置段落
- **50 / 72** 能回溯到 1965 年 Docket 4080 初版条文（另 22 条是后加条款，本来就没有）
- **15 / 72** 带更正件/校正提示，已在文档内以提示框标出
- 涉及 18 个修订案

**本轮修掉的取数层 bug**：

- `part23_source.mine_doc()` 去重时写成了 `b[0][:120]`（`b` 是 `(i, blk)` 元组，
  `b[0]` 是 int）→ 改 `b[1][0][:120]`。此 bug 让批量调用直接抛异常。
- 批量生成时新增 `_dedupe_key()` / `dedupe()`：同一段修订指令在 NPRM 与 Final Rule
  里各出现一次，需去重且**优先保留 Final Rule（颁行）版**，并剔除被截断的短块前缀。

用法：`python build_subpart_doc.py C`（或 B/D/E/F/G/A，`--limit N` 调试）。

### T7 — 铺开其余分部 ⬜（T6 脚本已可一键跑，见上表）

批量命令：

```bash
for s in B D E F G A; do python build_subpart_doc.py $s; done
```

B(52) → D(82) → E(87) → F(57) → G(31) → A(3) → 附录

### T8 — 总索引卷 ⬜

全局条款表 / 修订案表 / **文献反向索引**（哪份 NPRM、Final Rule 被哪些条款引用）

### T9 — 清理决策 ⬜ 待用户

`Rulemaking Docs/` 残留 **11 份早期 govinfo 路线 PDF**（命名形如 `FinalRule_Docket27805_…`，
Docket 后无空格）。其中至少 27805、28417 已确认被 DRS 版本覆盖。
**未自动删除**，等用户决定是否保留（可能想留作独立来源）。

---

## 五、环境要点与踩坑记录

### 数据源

- **FAA DRS**（drs.faa.gov）是 Angular SPA。已逆向端点：
  - `POST /api/browse/doctype/{label}/documents/metadatas`（body 加 **`"saveResult": true`**
    可突破 50 条上限，一次拿全集）
  - `GET /api/browse/documents/summary/{docId}?docTypeId=8`
  - `GET /api/browse/documents/summaryguid/{UNID大写}` ← UNID **必须大写**，小写 404
  - `GET /api/content/alf/{id}`
  - 主文档 PDF 时 `docContent` 为 null，要拿 `id` 走 `/api/content/alf/`；
    HTML 时 `docContent` 直接是 base64 全文
- **DRS 的 jwt 只有 12 小时**（iat→exp = 43200s）。2026-10-05 20:00 签发的，
  10-06 上午全部 `/api/*` 返回 403。**jwt 由前端 JS 写入，`curl -c` 拿不到新的**。
  → 源文档已全部落地，**尽量别再依赖 DRS**。
- **govinfo**：1994 年起有单篇 PDF（`digitizedFR: false`）；1994 年前只有整期影印
  （`digitizedFR: true`，100–250 MB）。**整期路线已废弃**。
- **eCFR versioner 对 Title 14 只覆盖 2016-12-30 之后**，历史版本查不到；
  但校对 Amdt 23-62 之后的现行条文很好用，⚠️ 必须带 `--compressed`（否则 406）。
- **Wayback / rgl.faa.gov 均不可用**（429 / 主机已下线）。
- 无头 Chrome 本机可跑，但必须 `--no-sandbox`，且**启动+使用+kill 要在同一条 Bash 调用内**。

### 文本处理

- 部分 txt 是 **CR(\r) 行尾**（如 Docket 25147），`sed`/`grep` 会把整篇当成一行
  → 必须用 Python universal newlines 读，先 `re.sub(r'[ \t]+',' ')` 再按"紧凑段序号"定位。
- **DRS 给的 FR 引证可能是文中某页而非首页**（Docket 26269 标 58 FR 42165，
  实际文档起于 58 FR 42136）→ 不能只按首页定位。
- 同一 DRS 文档可能同时出现在 FinalRule 和 NPRM 两个 doctype 里 → **去重必须按 UNID**。
- NPRM 元数据的 `Amendment` 字段常为空 → **不能反推 NPRM 覆盖哪些修订案**，
  必须用 `FAR23_条款-规则文件对应表.csv`。
- 文件名日期取自 DRS 的 `Issue Date`（签发日），**不是 FR 出版日**，两者常差数日至数周
  → 已用 `rename_with_fr_citation.py` 改为真实 FR 引证命名（150 组改名成功，11 组保留）。

### 工具链

- `python-docx` 装在 `/Users/glennchou/.workbuddy-ai/binaries/python/envs/default`
- `pypdf` / `pymupdf` 可用，`pdfplumber` / `PyPDF2` 无
- **macOS 无 LibreOffice**（`soffice` 是坏链接），无法转 PDF 预览版式
- risingup 页面缓存：`/tmp/risingup_pages`（381 个 .shtml，重跑无需联网）

---

## 六、脚本清单

| 脚本 | 作用 | 状态 |
|---|---|---|
| `enumerate_part23.py` | 枚举 DRS 条款清单 | 已完成 |
| `fetch_rulemaking.py` | 按 UNID 抓 NPRM/Final Rule 全文 | 已完成 |
| `fetch_from_universe.py` | 备选抓取通路 | 已完成 |
| `rename_with_fr_citation.py` | 用真实 FR 引证重命名（支持 `--apply`） | 已完成 |
| `coverage_report.py` | 62 个修订案 × FinalRule/更正件/NPRM 覆盖核算 | 可重跑 |
| `crawl_risingup.py` | 抓 381 页 + 解析溯源注 + 判定更正件 + 出两份 CSV | 已完成（本轮修 2 个 bug） |
| `part23_source.py` | 批量生成的数据层 | 已完成，待批量验证 |
| `build_23_561_doc.py` | §23.561 文档生成器（模板） | 已完成，待抽公共模块 |

---

## 七、当前状态与下一步

**T1–T6 全部完成**，无阻塞点。剩 T7（铺开其余 6 个分部，脚本已就绪）、
T8（总索引卷）、T9（`Rulemaking Docs/` 里 11 份早期 govinfo PDF 是否清理，等用户决定）。

可立即执行：

1. `for s in B D E F G A; do python build_subpart_doc.py $s; done` —— 铺开其余分部
2. 需要哪几条条款的**精写版**（含彩色条文 diff），指定条款号即可按 §23.561 模板做
