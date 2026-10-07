#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成《14 CFR § 23.561 条款修订历史与背景分析报告》Word 文档。

数据源（全部为本地已下载文件，不再联网）：
  - FAR23_Amdt23-63_条款清单.csv      DRS 条款修订版记录（日期 / FR 案号）
  - Rulemaking Docs/*                 NPRM 与 Final Rule 官方原文（FR 影印件 + 提取文本）
  - ecfr.gov 2017-01-01 point-in-time XML（用于校对 Amdt 23-62 之后的条文）

着色约定：
  蓝 = 本版新增   红 = 本版修订/替换   灰+删除线 = 本版删除   黑 = 继承未变
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "/Users/glennchou/FAR 23/FAR23_Sec23.561_修订历史与背景分析.docx"

# ---------------- 颜色 / 字体 ----------------
C_NEW = RGBColor(0x00, 0x70, 0xC0)      # 新增：蓝
C_REV = RGBColor(0xC0, 0x00, 0x00)      # 修订：红
C_DEL = RGBColor(0x80, 0x80, 0x80)      # 删除：灰
C_TXT = RGBColor(0x1A, 0x1A, 0x1A)      # 正文
C_HDR_BG = "1F4E79"                     # 表头底色
C_SUB_BG = "DDEBF7"                     # 交替/次级底色
C_NOTE_BG = "FFF2CC"                    # 提示框底色
C_BOX_BG = "F2F2F2"                     # 条文框底色
CN_FONT = "微软雅黑"
EN_FONT = "Calibri"


# ---------------- 底层工具 ----------------
def set_run(run, size=10.5, bold=False, color=C_TXT, italic=False,
            strike=False, cn=CN_FONT, en=EN_FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = en
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    if strike:
        run.font.strike = True
    return run


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def cell_margins(table, top=60, bottom=60, left=90, right=90):
    tblPr = table._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for tag, val in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        e = OxmlElement('w:' + tag)
        e.set(qn('w:w'), str(val))
        e.set(qn('w:type'), 'dxa')
        mar.append(e)
    tblPr.append(mar)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement('w:tblHeader')
    e.set(qn('w:val'), 'true')
    trPr.append(e)


def fill_runs(p, segs, size):
    """把 segs 渲染进段落 p。segs 可为 str、[(text, style)] 或二者的混合列表。"""
    if isinstance(segs, str):
        segs = [(segs, 'normal')]
    for item in segs:
        if isinstance(item, str):
            text, st = item, 'normal'
        else:
            text, st = item
        r = p.add_run(text)
        if st == 'new':
            set_run(r, size=size, bold=True, color=C_NEW)
        elif st == 'rev':
            set_run(r, size=size, bold=True, color=C_REV)
        elif st == 'del':
            set_run(r, size=size, color=C_DEL, strike=True)
        elif st == 'bold':
            set_run(r, size=size, bold=True, color=C_TXT)
        elif st == 'head':
            set_run(r, size=size, bold=True, color=RGBColor(0x1F, 0x4E, 0x79))
        else:
            set_run(r, size=size, color=C_TXT)
    return p


def para_spacing(p, before=4, after=4, line=1.35):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def add_toc_field(doc):
    p = doc.add_paragraph()
    r = p.add_run()
    for typ, txt in (('begin', None), ('separate', "　※ 在 Word/WPS 中按 Ctrl+A 后 F9（或右键→更新域）生成目录"),
                     ('end', None)):
        f = OxmlElement('w:fldChar')
        f.set(qn('w:fldCharType'), typ)
        r._r.append(f)
        if typ == 'begin':
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = r'TOC \o "1-2" \h \z \u'
            r._r.append(it)
        elif typ == 'separate':
            t = OxmlElement('w:t')
            t.text = txt
            r._r.append(t)
    set_run(r, size=10, italic=True, color=C_DEL)
    return p


def add_page_number_footer(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("14 CFR § 23.561 修订历史与背景分析　|　第 ")
    set_run(r1, size=8.5, color=C_DEL)
    r = p.add_run()
    f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'begin'); r._r.append(f)
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = 'PAGE'; r._r.append(it)
    f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'end'); r._r.append(f)
    set_run(r, size=8.5, color=C_DEL)
    r2 = p.add_run(" 页")
    set_run(r2, size=8.5, color=C_DEL)


# ---------------- 高层组件 ----------------
def H(doc, text, level=1, size=None, color=None, before=14, after=6):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = True
    r = p.add_run(text)
    size = size or {1: 16, 2: 13.5, 3: 12}[level]
    color = color or {1: RGBColor(0x1F, 0x4E, 0x79),
                      2: RGBColor(0x2E, 0x74, 0xB5),
                      3: RGBColor(0x40, 0x40, 0x40)}[level]
    set_run(r, size=size, bold=True, color=color)
    p.style = doc.styles['Normal']
    return p


def P(doc, text, size=10.5, indent=0, bold=False, color=C_TXT, before=3, after=3,
      italic=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    para_spacing(p, before, after)
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=color, italic=italic)
    return p


def RICH(doc, segments, size=10.5, indent=0, before=3, after=3):
    """segments: [(text, style)]  style in normal/new/rev/del/bold"""
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    para_spacing(p, before, after)
    for text, st in segments:
        r = p.add_run(text)
        if st == 'new':
            set_run(r, size=size, bold=True, color=C_NEW)
        elif st == 'rev':
            set_run(r, size=size, bold=True, color=C_REV)
        elif st == 'del':
            set_run(r, size=size, color=C_DEL, strike=True)
        elif st == 'bold':
            set_run(r, size=size, bold=True, color=C_TXT)
        else:
            set_run(r, size=size, color=C_TXT)
    return p


def BULLET(doc, label, segments, indent=0.6, size=10.5):
    """label 为加粗小标题，segments 为 [(text, style)]"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.first_line_indent = Cm(-0.45)
    para_spacing(p, 3, 3)
    r = p.add_run("▪ " + label)
    set_run(r, size=size, bold=True, color=RGBColor(0x1F, 0x4E, 0x79))
    for text, st in segments:
        rr = p.add_run(text)
        if st == 'new':
            set_run(rr, size=size, bold=True, color=C_NEW)
        elif st == 'rev':
            set_run(rr, size=size, bold=True, color=C_REV)
        elif st == 'del':
            set_run(rr, size=size, color=C_DEL, strike=True)
        else:
            set_run(rr, size=size, color=C_TXT)
    return p


def mk_table(doc, headers, rows, widths=None, font=9.5, header_bg=C_HDR_BG,
             zebra=True, align_center_cols=()):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    cell_margins(t)
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para_spacing(p, 2, 2)
        set_run(p.add_run(h), size=font, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(c, header_bg)
    repeat_header(hdr)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]
            c.text = ""
            p = c.paragraphs[0]
            para_spacing(p, 2, 2)
            if i in align_center_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fill_runs(p, val, font)
            if zebra and ri % 2 == 1:
                shade(c, "F7F9FC")
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t


def info_box(doc, title, lines, bg=C_NOTE_BG):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    cell_margins(t)
    c = t.rows[0].cells[0]
    c.text = ""
    shade(c, bg)
    p = c.paragraphs[0]
    para_spacing(p, 2, 3)
    set_run(p.add_run(title), size=10.5, bold=True, color=RGBColor(0x1F, 0x4E, 0x79))
    for ln in lines:
        pp = c.add_paragraph()
        para_spacing(pp, 1, 1)
        if isinstance(ln, list):
            # [(text, style), ...]
            for text, st in ln:
                r = pp.add_run(text)
                if st == 'bold':
                    set_run(r, size=9.5, bold=True, color=C_TXT)
                elif st == 'new':
                    set_run(r, size=9.5, bold=True, color=C_NEW)
                elif st == 'rev':
                    set_run(r, size=9.5, bold=True, color=C_REV)
                elif st == 'del':
                    set_run(r, size=9.5, color=C_DEL, strike=True)
                else:
                    set_run(r, size=9.5, color=C_TXT)
        else:
            set_run(pp.add_run("　" + ln), size=9.5, color=C_TXT)
    doc.add_paragraph()
    return t


def legend(doc):
    p = doc.add_paragraph()
    para_spacing(p, 2, 6)
    for text, col, bold in (("■ 蓝色", C_NEW, True), ("＝本版新增　　", C_TXT, False),
                            ("■ 红色", C_REV, True), ("＝本版修订/替换　　", C_TXT, False),
                            ("■ 灰色删除线", C_DEL, True), ("＝本版删除　　", C_TXT, False),
                            ("■ 黑色", C_TXT, False), ("＝继承上版、未变化", C_TXT, False)):
        r = p.add_run(text)
        set_run(r, size=9.5, bold=bold, color=col, strike=(col == C_DEL and bold))
    return p


def clause_table(doc, title, rows, note=None):
    """rows: [(段落号, [(文本, style)], 状态标记)]"""
    P(doc, title, size=11, bold=True, color=RGBColor(0x1F, 0x4E, 0x79), before=10, after=3)
    if note:
        P(doc, note, size=9.5, italic=True, color=C_DEL, indent=0.3, before=0, after=4)
    t = doc.add_table(rows=1, cols=3)
    t.style = 'Table Grid'
    t.autofit = False
    cell_margins(t)
    hdr = t.rows[0]
    for i, h in enumerate(["段落", "条文内容（英文原文）", "状态"]):
        c = hdr.cells[i]
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para_spacing(p, 2, 2)
        set_run(p.add_run(h), size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(c, C_HDR_BG)
    repeat_header(hdr)
    for label, segs, status in rows:
        cells = t.add_row().cells
        # 段落号
        cells[0].text = ""
        p = cells[0].paragraphs[0]
        para_spacing(p, 2, 2)
        set_run(p.add_run(label), size=9.5, bold=True)
        shade(cells[0], C_BOX_BG)
        # 条文
        cells[1].text = ""
        p = cells[1].paragraphs[0]
        para_spacing(p, 2, 2)
        for text, st in segs:
            r = p.add_run(text)
            if st == 'new':
                set_run(r, size=9.5, bold=True, color=C_NEW)
            elif st == 'rev':
                set_run(r, size=9.5, bold=True, color=C_REV)
            elif st == 'del':
                set_run(r, size=9.5, color=C_DEL, strike=True)
            elif st == 'head':
                set_run(r, size=9.5, bold=True, color=RGBColor(0x1F, 0x4E, 0x79))
            else:
                set_run(r, size=9.5, color=C_TXT)
        # 状态
        cells[2].text = ""
        p = cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para_spacing(p, 2, 2)
        col = {"新增": C_NEW, "修订": C_REV, "删除": C_DEL}.get(status, C_TXT)
        set_run(p.add_run(status), size=9, bold=(status != "—"), color=col)
        if status == "新增":
            shade(cells[2], "E7F1FA")
        elif status == "修订":
            shade(cells[2], "FBE9E9")
        elif status == "删除":
            shade(cells[2], "F0F0F0")
    for i, w in enumerate((1.6, 13.6, 1.6)):
        for row in t.rows:
            row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t


# =====================================================================
# 数据：7 个版本的条文
# =====================================================================
HEAD = ("Sec. 23.561 General.", 'head')

V1_ROWS = [
    ("(a)", [("The airplane, although it may be damaged in emergency landing conditions, must be "
              "designed as prescribed in this section to protect each occupant under those conditions.", 'normal')], "—"),
    ("(b)", [("The structure must be designed to give each occupant every reasonable chance of escaping "
              "serious injury in a minor crash landing when—", 'normal')], "—"),
    ("(b)(1)", [("Proper use is made of belts or harnesses provided for in the design; and", 'normal')], "—"),
    ("(b)(2)", [("The occupant experiences the ultimate inertia forces shown in the following table:  "
                 "［表：Ultimate Inertia Forces］Upward 3.0g（正常/实用类）/ 4.5g（特技类）；"
                 "Forward 9.0g / 9.0g；Sideward 1.5g / 1.5g", 'normal')], "—"),
    ("(c)", [("Each airplane with retractable landing gear must be designed to protect each occupant in "
              "a landing—", 'normal')], "—"),
    ("(c)(1)", [("With the wheels retracted;", 'normal')], "—"),
    ("(c)(2)", [("With moderate descent velocity; and", 'normal')], "—"),
    ("(c)(3)", [("Assuming—", 'normal')], "—"),
    ("(c)(3)(i)", [("An upward ultimate inertia force of 3 g; and", 'normal')], "—"),
    ("(c)(3)(ii)", [("A coefficient of friction of 0.5 at the ground.", 'normal')], "—"),
    ("(d)", [("If a turnover is reasonably probable, the structure must be designed to protect the "
              "occupants in a complete turnover, assuming—", 'normal')], "—"),
    ("(d)(1)", [("An upward ultimate inertia force of 3 g; and", 'normal')], "—"),
    ("(d)(2)", [("A coefficient of friction of 0.5 at the ground.", 'normal')], "—"),
    ("(e)", [("Except as provided in Sec. 23.787 the supporting structure must be designed to restrain, "
              "under loads up to those specified in paragraph (b)(2) of this section, each item of mass "
              "that could injure an occupant if it came loose in a minor crash landing.", 'normal')], "—"),
]

V2_ROWS = [
    ("(a)", [("（同 Amdt 23-0 版，未变）", 'normal')], "—"),
    ("(b)", [("（同 Amdt 23-0 版，未变）", 'normal')], "—"),
    ("(c)(3)", [("Assuming, ", 'normal'), ("in the absence of a more rational analysis —", 'new'), ("", 'normal')], "修订"),
    ("(c)(3)(i)", [("A ", 'normal'), ("downward", 'rev'), (" ultimate inertia force of 3 g; and  ", 'normal'),
                   ("（原为 upward）", 'del')], "修订"),
    ("(c)(3)(ii)", [("A coefficient of friction of 0.5 at the ground.", 'normal')], "—"),
    ("(d)", [("If a turnover is reasonably probable, the structure must be designed to protect the "
              "occupants in a complete turnover, assuming, ", 'normal'),
             ("in the absence of a more rational analysis —", 'new')], "修订"),
    ("(d)(1)", [("An upward ultimate inertia force of 3 g; and", 'normal')], "—"),
    ("(d)(2)", [("A coefficient of friction of 0.5 at the ground.", 'normal')], "—"),
    ("(e)", [("（同 Amdt 23-0 版，未变）", 'normal')], "—"),
]

V4_ROWS = [
    ("(a)", [("（同前，未变）", 'normal')], "—"),
    ("(b)", [("The structure must be designed to protect each occupant during emergency landing "
              "conditions when:", 'rev')], "修订"),
    ("(b)(1)", [("Proper use is made of seats, safety belts, and shoulder harnesses provided for in "
                 "the design;", 'rev')], "修订"),
    ("(b)(2)", [("The occupant experiences the static inertia loads corresponding to the following "
                 "ultimate load factors—", 'rev')], "修订"),
    ("(b)(2)(i)", [("Upward, 3.0g for normal, utility, and commuter category airplanes, or 4.5g for "
                    "acrobatic category airplanes;", 'rev')], "修订"),
    ("(b)(2)(ii)", [("Forward, 9.0g;", 'rev')], "修订"),
    ("(b)(2)(iii)", [("Sideward, 1.5g; and", 'rev')], "修订"),
    ("(b)(3)", [("The items of mass within the cabin, that could injure an occupant, experience the "
                 "static inertia loads corresponding to the following ultimate load factors—", 'new')], "新增"),
    ("(b)(3)(i)", [("Upward, 3.0g;", 'new')], "新增"),
    ("(b)(3)(ii)", [("Forward, 18.0g; and", 'new')], "新增"),
    ("(b)(3)(iii)", [("Sideward, 4.5g", 'new')], "新增"),
    ("(c)", [("（Amdt 23-7 版 (c) 段保留：可收放起落架机轮收上着陆）", 'normal')], "—"),
    ("(d)", [("If it is not established that a turnover is unlikely during an emergency landing, the "
              "structure must be designed to protect the occupants in a complete turnover as follows:", 'rev')], "修订"),
    ("(d)(1)", [("The likelihood of a turnover may be shown by an analysis assuming the following "
                 "conditions—", 'rev')], "修订"),
    ("(d)(1)(i)", [("Maximum weight;", 'rev')], "修订"),
    ("(d)(1)(ii)", [("Most forward center of gravity position;", 'rev')], "修订"),
    ("(d)(1)(iii)", [("Longitudinal load factor of 9.0g;", 'rev')], "修订"),
    ("(d)(1)(iv)", [("Vertical load factor of 1.0g; and", 'rev')], "修订"),
    ("(d)(1)(v)", [("For airplanes with tricycle landing gear, the nose wheel strut failed with the "
                    "nose contacting the ground.", 'rev')], "修订"),
    ("(d)(2)", [("For determining the loads to be applied to the inverted airplane after a turnover, an "
                 "upward ultimate inertia load factor of 3.0g and a coefficient of friction with the "
                 "ground of 0.5 must be used.", 'rev')], "修订"),
    ("(e)", [("Except as provided in Sec. 23.787 the supporting structure must be designed to restrain, "
              "under loads up to those specified in paragraph (b)(2) of this section, each item of mass "
              "that could injure an occupant if it came loose in a minor crash landing.", 'del'),
             ("　［本版重组后本段不再独立存在，其载荷要求并入 (b)(3)；1996 年 Amdt 23-48 以“新增 (e)”方式恢复］", 'normal')], "删除"),
]

V5_ROWS = [
    ("(b)(2)(iv)", [("Downward, 6.0g when certification to the emergency exit provisions of "
                     "Sec. 23.807(d)(4) is requested; and", 'new')], "新增"),
    ("其余", [("（同 Amdt 23-36 版，未变）", 'normal')], "—"),
]

V6_ROWS = [
    ("(b)", [("The structure must be designed to give each occupant every reasonable chance of escaping "
              "serious injury when—", 'rev'),
             ("　［1988 版为 “…to protect each occupant during emergency landing conditions when:”］", 'normal')], "修订"),
    ("(d)(1)", [("The likelihood of a turnover may be shown by an analysis assuming the following "
                 "conditions—", 'normal')], "—"),
    ("(d)(1)(i)", [("The most adverse combination of weight and center of gravity position;", 'rev'),
                   ("　［合并 1988 版 (i) 最大重量 与 (ii) 最前重心］", 'normal')], "修订"),
    ("(d)(1)(ii)", [("Longitudinal load factor of 9.0g;", 'rev')], "修订"),
    ("(d)(1)(iii)", [("Vertical load factor of 1.0g; and", 'rev')], "修订"),
    ("(d)(1)(iv)", [("For airplanes with tricycle landing gear, the nose wheel strut failed with the "
                     "nose contacting the ground.", 'rev')], "修订"),
    ("(d)(1)(v)", [("Maximum weight.", 'del'), ("　［已并入新的 (d)(1)(i)］", 'normal')], "删除"),
    ("(d)(2)", [("（同 Amdt 23-36 版，未变）", 'normal')], "—"),
    ("(e)", [("Except as provided in Sec. 23.787(c), the supporting structure must be designed to "
              "restrain, under loads up to those specified in paragraph (b)(3) of this section, each "
              "item of mass that could injure an occupant if it came loose in a minor crash landing.", 'new'),
             ("　［相比初始版：引用由 Sec. 23.787 精确到 Sec. 23.787(c)，载荷引用由 (b)(2) 更新为 (b)(3)］", 'normal')], "新增"),
]

V7_ROWS = [
    ("(e)", [("Except as provided in Sec. 23.787(c), the supporting structure must be designed to "
              "restrain, under loads up to those specified in paragraph (b)(3) of this section, each "
              "item of mass that could injure an occupant if it came loose in a minor crash landing.", 'normal')], "—"),
    ("(e)(1)", [("For engines mounted inside the fuselage, aft of the cabin, it must be shown by test or "
                 "analysis that the engine and attached accessories, and the engine mounting structure—", 'new')], "新增"),
    ("(e)(1)(i)", [("Can withstand a forward acting static ultimate inertia load factor of 18.0 g plus "
                    "the maximum takeoff engine thrust; or", 'new')], "新增"),
    ("(e)(1)(ii)", [("The airplane structure is designed to preclude the engine and its attached "
                     "accessories from entering or protruding into the cabin should the engine mounts "
                     "fail.", 'new'),
                    ("　［NPRM 原作 “to deflect the engine … away from the cabin”；"
                     "采纳 Transport Canada 意见后改为性能化表述］", 'normal')], "新增"),
    ("(e)(2)", [("[Reserved]", 'new')], "新增"),
    ("其余", [("（同 Amdt 23-48 版，未变）", 'normal')], "—"),
]

# 关键要素演变矩阵
MATRIX_HEADERS = ["条款要素", "23-0\n(1965)", "23-7\n(1969)", "23-34\n(1987)", "23-36\n(1988)",
                  "23-46\n(1994)", "23-48\n(1996)", "23-62\n(2012)"]
MATRIX_ROWS = [
    ["设计场景表述", "minor crash landing", "同左", "同左",
     [("emergency landing conditions", 'rev')], "同左", "同左", "同左"],
    ["乘员 Upward", "3.0g / 4.5g(特技)", "同左", "同左",
     [("3.0g (正常·实用·通勤) / 4.5g (特技)", 'rev')], "同左", "同左", "同左"],
    ["乘员 Forward", "9.0g", "同左", "同左", "9.0g", "同左", "同左", "同左"],
    ["乘员 Sideward", "1.5g", "同左", "同左", "1.5g", "同左", "同左", "同左"],
    ["乘员 Downward", "—", "—", "—", "—", [("6.0g（按 §23.807(d)(4) 取证时）", 'new')], "同左", "同左"],
    ["质量块 Upward", "随 (b)(2) 3.0g", "同左", "同左", [("3.0g　(b)(3)(i)", 'new')], "同左", "同左", "同左"],
    ["质量块 Forward", "随 (b)(2) 9.0g", "同左", "同左", [("18.0g　(b)(3)(ii)", 'new')], "同左", "同左", "同左"],
    ["质量块 Sideward", "随 (b)(2) 1.5g", "同左", "同左", [("4.5g　(b)(3)(iii)", 'new')], "同左", "同左", "同左"],
    ["机轮收上着陆 (c)(3)(i)", [("upward", 'del'), " 3g"], [("downward", 'rev'), " 3g"], "同左", "同左", "同左", "同左", "downward 3 g"],
    ["“合理分析”许可", "无", [("加入 in the absence of a more rational analysis", 'new')], "同左", "同左", "同左", "同左", "同左"],
    ["翻转判据", [("If a turnover is reasonably probable", 'del')], "同左", "同左",
     [("If it is not established that a turnover is unlikely", 'rev')], "同左", "同左", "同左"],
    ["翻转分析条件数", "—", "—", "—", [("5 项", 'rev')], "同左", [("4 项（合并重量与重心）", 'rev')], "同左"],
    ["(e) 支撑结构约束质量块", "有（引用 §23.787、(b)(2)）", "同左", "同左",
     [("无（并入 (b)(3)）", 'del')], "同左", [("有（引用 §23.787(c)、(b)(3)）", 'new')], "有"],
    ["机身内发动机坠撞约束", "—", "—", "—", "—", "—", "—",
     [("(e)(1) 18.0g＋最大起飞推力，或阻止侵入客舱", 'new')]],
    ["适用飞机类别", "正常/实用/特技", "同左", "+通勤类（类目新增）", "正常/实用/通勤/特技", "同左", "同左", "+通勤类喷气机"],
]

# 修订沿革总表
HIST_HEADERS = ["#", "修订号", "生效日期", "NPRM（编号 / 签发 / FR 引证）",
                "Final Rule（案号 / 签发 / FR 引证）", "本版对 §23.561 的修订内容"]
HIST_ROWS = [
    ["1", "Amdt 23-0\n（初始版）", "1965-02-01",
     "Notice 64-17\nDocket 4080\n1964-03-25 签发\n29 FR 5111\n(1964-04-14)",
     "Docket 4080\n1964-09-28 签发\n29 FR 17955\n(1964-12-18)",
     "由 CAR 3.386 重编为 FAR §23.561（NPRM 中编号 §23.709 “Protection”），确立 (a)–(e) 五段结构；"
     "给出乘员极限静惯性力表 Upward 3.0g/4.5g、Forward 9.0g、Sideward 1.5g；(c) 机轮收上着陆；"
     "(d) 翻转；(e) 支撑结构约束质量块。"],
    ["2", "Amdt 23-7", "1969-09-14",
     "Notice 67-14\nDocket 8083\n1967-04-05 签发\n32 FR 5791\n(1967-04-11)",
     "Docket 8083\n1969-08-01 签发\n34 FR 13078\n(1969-08-13)",
     "(c)(3) 与 (d) 引导句加入 “in the absence of a more rational analysis”；"
     "(c)(3)(i) 由 upward 更正为 downward。澄清性修改，不改变数值。"],
    ["3", "Amdt 23-34", "1987-02-17",
     "Notice 83-17\nDocket 23516\n1983-09-23 签发\n48 FR 52010\n(1983-11-15)",
     "Docket 23516\n1987-01-08 签发\n52 FR 1806\n(1987-01-15)",
     "条文未改动（52 FR 1806 共 43 条修订指令，不含 §23.561）。本版新增通勤类（Commuter Category），"
     "§23.561 适用范围随之扩展，并被 §23.785(g)、§23.807(d)、§23.963(f) 等新增条款交叉引用。"],
    ["4", "Amdt 23-36", "1988-09-14",
     "Notice 86-19\nDocket 25147\n1986-11-28 签发\n51 FR 44878\n(1986-12-12)",
     "Docket 25147\n1988-08-08 签发\n53 FR 30802\n(1988-08-15)",
     "(b) 全面重写：改用 “emergency landing conditions” 表述、载荷表按飞机类别重写、"
     "新增 (b)(3) 客舱质量块 18.0g/4.5g；(d) 翻转判据客观化（5 项分析条件 + 倒置载荷）；"
     "配套新增 §23.562 动态着陆条件，形成“§23.561 静态 / §23.562 动态”分工。"],
    ["5", "Amdt 23-46", "1994-06-16",
     "Notice 90-20\nDocket 26324\n1990-08-23 签发\n55 FR 35544\n(1990-08-30)",
     "Docket 26324\n1994-05-11 签发\n59 FR 25766\n(1994-05-17)",
     "新增 (b)(2)(iv)：Downward 6.0g（当申请人选择按 §23.807(d)(4) 替代应急出口取证时）。"
     "把原规定在 §23.807(d)(1) 豁免条款中的向下载荷移入 §23.561，与 Part 25 统一。"],
    ["6", "Amdt 23-48", "1996-03-11",
     "Notice 94-20\nDocket 27805\n1994-07-05 签发\n59 FR 35196\n(1994-07-08)",
     "Docket 27805\n1996-01-29 签发\n61 FR 5138\n(1996-02-09)",
     "FAR/JAR 协调：(b) 引导句改为与 Part 25/JAR 25 一致；(d)(1) 由 5 项精简为 4 项；"
     "新增 (e) 支撑结构约束质量块（引用更新为 §23.787(c)、(b)(3)）。"],
    ["7", "Amdt 23-62", "2012-01-31",
     "Docket FAA-2009-0738\n（本时期已不使用 Notice 号）\n2009-08-06 签发\n74 FR 41522\n(2009-08-17)",
     "Docket FAA-2009-0738\n2011-12-02 签发\n76 FR 75736\n(2011-12-02)",
     "新增 (e)(1)：机身内、客舱后方的发动机须承受前向 18.0g + 最大起飞推力，"
     "或机体结构设计为在发动机架失效时阻止发动机侵入客舱；(e)(2) [Reserved]。"
     "同期通勤类扩展至喷气机，并修订 §23.562。"],
]

# 参考文献
REF_ROWS = [
    ["[1]", "NPRM", "Notice 64-17 / Docket No. 4080",
     "29 FR 5111，1964-04-14（签发 1964-03-25）",
     "NPRM_Notice 64-17_Docket4080_29FR5111_1964-04-14.pdf / .txt"],
    ["[2]", "Final Rule", "Docket No. 4080（Amdt 23-0）",
     "29 FR 17955，1964-12-18（签发 1964-09-28；生效 1965-02-01）",
     "FinalRule_Docket 4080_Amdt23-0_1964-12-18.pdf / _0002.txt"],
    ["[3]", "NPRM", "Notice 67-14 / Docket No. 8083",
     "32 FR 5791，1967-04-11（签发 1967-04-05）",
     "NPRM_Notice 67-14_Docket8083_32FR5791_1967-04-11.pdf / .txt"],
    ["[4]", "Final Rule", "Docket No. 8083（Amdt 23-7）",
     "34 FR 13078，1969-08-13（签发 1969-08-01；生效 1969-09-14）",
     "FinalRule_Docket 8083_Amdt23-7_34FR13078_1969-08-13.pdf / _0002.txt"],
    ["[5]", "NPRM", "Notice 83-17 / Docket No. 23516",
     "48 FR 52010，1983-11-15（签发 1983-09-23）",
     "NPRM_Notice 83-17_Docket23516_48FR52010_1983-11-15.pdf / .txt"],
    ["[6]", "Final Rule", "Docket No. 23516（Amdt 23-34，含 21-59/36-13/91-197/135-21）",
     "52 FR 1806，1987-01-15（签发 1987-01-08；生效 1987-02-17）",
     "FinalRule_Docket 23516_Amdt21-59_23-34_36-13_52FR_1987-01-15.html / .txt"],
    ["[7]", "NPRM", "Notice 86-19 / Docket No. 25147",
     "51 FR 44878，1986-12-12（签发 1986-11-28）",
     "NPRM_Notice 86-19_Docket25147_51FR44878_1986-12-12.pdf / .txt"],
    ["[8]", "Final Rule", "Docket No. 25147（Amdt 23-36）",
     "53 FR 30802，1988-08-15（签发 1988-08-08；生效 1988-09-14）",
     "FinalRule_Docket 25147_Amdt23-36_53FR30802_1988-08-15.pdf / _0002.txt"],
    ["[9]", "NPRM", "Notice 90-20 / Docket No. 26324",
     "55 FR 35544，1990-08-30（签发 1990-08-23）",
     "NPRM_Notice 90-20_Docket26324_55FR35544_1990-08-30.pdf / .txt"],
    ["[10]", "Final Rule", "Docket No. 26324（Amdt 23-46）",
     "59 FR 25766，1994-05-17（签发 1994-05-11；生效 1994-06-16）",
     "FinalRule_Docket 26324_Amdt23-46_59FR25766_1994-05-17.pdf / _0002.txt"],
    ["[11]", "NPRM", "Notice 94-20 / Docket No. 27805",
     "59 FR 35196，1994-07-08（签发 1994-07-05）",
     "NPRM_Notice 94-20_Docket27805_59FR35196_1994-07-08.pdf / .txt"],
    ["[12]", "Final Rule", "Docket No. 27805（Amdt 23-48）",
     "61 FR 5138，1996-02-09（签发 1996-01-29；生效 1996-03-11）",
     "FinalRule_Docket 27805_Amdt23-48_61FR5138_1996-02-09.pdf / _0002.txt"],
    ["[13]", "NPRM", "Docket No. FAA-2009-0738（无 Notice 号）",
     "74 FR 41522，2009-08-17（签发 2009-08-06）",
     "NPRM_Docket FAA-2009-0738_74FR41522_2009-08-17.pdf"],
    ["[14]", "Final Rule", "Docket No. FAA-2009-0738（Amdt 23-62）",
     "76 FR 75736，2011-12-02（生效 2012-01-31）",
     "FinalRule_Docket FAA-2009-0738_Amdt23-62_76FR75736_2011-12-02.html / .txt"],
    ["[15]", "条文校对", "eCFR point-in-time XML（2017-01-01 版）",
     "14 CFR 23.561，用于核对 Amdt 23-62 之后的完整条文",
     "ecfr.gov/api/versioner/v1/full/2017-01-01/title-14.xml?part=23&section=23.561"],
    ["[16]", "修订记录", "FAA 动态法规系统（DRS）条款修订版索引",
     "Part 23 @ Amdt 23-63 as of 08/29/2017 and earlier",
     "FAR23_Amdt23-63_条款清单.csv（本地）"],
    ["[17]", "更正通告", "Docket No. 23516（Amdt 21-59/23-34/36-13/91-197/135-21 更正）",
     "52 FR 34745，1987-09-14",
     "Corrections/Correction_52FR34745_1987-09-14.clean.txt（govinfo 整期影印 PDF p.137 抽取）"],
    ["[18]", "CFR 编纂校正", "OFR CFR Correction（14 CFR Part 23 §23.561）",
     "72 FR 59939，2007-10-23",
     "Corrections/CFRCorrection_72FR59939_2007-10-23_07-55519.clean.txt"],
    ["[19]", "CFR 编纂校正", "OFR CFR Correction（上一件的重发修正）",
     "72 FR 72915，2007-12-26",
     "Corrections/CFRCorrection_72FR72915_2007-12-26_07-55522.clean.txt"],
]

DRS_ROWS = [
    ["23-0（Initial）", "1965-02-01", "c1b8f6a9-56f2-41c5-8f4e-5847a836a3ee"],
    ["23-7", "1969-09-14", "70523e6c-650a-4940-95d9-d0c5777c0300"],
    ["23-34", "1987-02-17", "564de8fe-3467-4abb-b430-266d540577e0"],
    ["23-36", "1988-09-14", "024246fb-52f0-4063-a39a-19e8642ee44d"],
    ["23-46", "1994-06-16", "1d488d69-dc82-4181-9237-96857e36471c"],
    ["23-48", "1996-03-11", "4e7946eb-090b-4845-81db-697548c2aa83"],
    ["23-62", "2012-01-31", "e3bb01c4-2726-4f09-bf08-7af03e7fae95"],
]


# =====================================================================
# 组装文档
# =====================================================================
def build():
    doc = Document()

    # 页面设置
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.0)

    # 默认字体
    st = doc.styles['Normal']
    st.font.name = EN_FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), CN_FONT)
    st.paragraph_format.line_spacing = 1.35

    add_page_number_footer(sec)

    # ---------- 封面 ----------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, 40, 4)
    set_run(p.add_run("14 CFR Part 23 条款修订历史与背景分析报告"),
            size=22, bold=True, color=RGBColor(0x1F, 0x4E, 0x79))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, 0, 6)
    set_run(p.add_run("§ 23.561　General.（应急着陆通用条件）"),
            size=16, bold=True, color=RGBColor(0x2E, 0x74, 0xB5))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, 0, 18)
    set_run(p.add_run("1965 年 2 月 1 日 初版 ｜ 2012 年 1 月 31 日 Amdt 23-62 ｜ 共 7 个历史版本"),
            size=11, color=RGBColor(0x59, 0x59, 0x59))

    info_box(doc, "文档说明与编制依据", [
        [("• 修订日期、NPRM 编号、Final Rule 编号与日期：取自 ", 'normal'),
         ("FAA 动态法规系统（DRS）", 'bold'),
         ("「Part 23 @ Amdt 23-63 as of 08/29/2017 and earlier」条款修订版索引（见附录）。", 'normal')],
        "• 修订背景、原因、评论处置与条文原文：逐条引自本地已下载的 NPRM 与 Final Rule 官方文本"
        "（Federal Register 影印件及提取文本），并在第 5 节参考文献中逐条标明 FR 卷页、日期与本地文件名。",
        "• 2012 版（Amdt 23-62）条文以 eCFR point-in-time XML（2017-01-01）交叉校对。",
        "• 文中所有引用均以 [n] 标注，对应第 5 节参考文献编号；条文对比中的颜色含义见第 3 节图例。",
        "• 本报告范围为 Amdt 23-63 及更早的历史版本；2017 年 Amdt 23-64 的性能化重构不在本快照内，"
        "仅在结语中作提示。",
    ])

    # ---------- 目录 ----------
    H(doc, "目录", 1)
    add_toc_field(doc)
    for ln in ["1　§23.561 修订历史汇总",
               "2　各阶段修订背景与原因深度剖析",
               "　　2.1　初始颁行（1964–1965，Amdt 23-0）：从 CAR 3.386 到 FAR §23.561",
               "　　2.2　Amdt 23-7（1969）：两处澄清性更正",
               "　　2.3　Amdt 23-34（1987）：通勤类引入与“未改之改”",
               "　　　　2.3.1　考据：CFR 溯源注为何把 52 FR 34745 标成 Amdt 23-24",
               "　　2.4　Amdt 23-36（1988）：坠撞动力学与静态/动态分工",
               "　　2.5　Amdt 23-46（1994）：替代应急出口与向下 6.0g",
               "　　2.6　Amdt 23-48（1996）：FAR/JAR 协调统一",
               "　　2.7　Amdt 23-62（2012）：机身内嵌发动机与性能化转向",
               "　　2.8　三份被溯源注遗漏的“更正”类文献（1987 / 2007）",
               "3　条款原文演变对比",
               "4　总结：§23.561 四十余年的演进规律",
               "5　参考文献",
               "附录　DRS 条款版本记录 ID"]:
        pp = doc.add_paragraph()
        pp.paragraph_format.left_indent = Cm(0.5 if ln.startswith("　") else 0)
        para_spacing(pp, 1, 1)
        set_run(pp.add_run(ln), size=10,
                bold=not ln.startswith("　"),
                color=RGBColor(0x1F, 0x4E, 0x79) if not ln.startswith("　") else C_TXT)

    doc.add_page_break()

    # ================= 1. 修订历史汇总 =================
    H(doc, "1　§23.561 修订历史汇总", 1)

    H(doc, "1.1　条款基本信息", 2)
    mk_table(doc, ["项目", "内容"], [
        ["条款号", "14 CFR § 23.561"],
        ["条款标题", "General.（应急着陆通用条件）"],
        ["所属分部", "Subpart C — Structure（C 分部·结构）"],
        ["适用类别", "正常类 / 实用类 / 特技类 / 通勤类（Normal, Utility, Acrobatic, Commuter Category）"],
        ["核心要求", "应急着陆条件下对乘员的保护：静惯性载荷因子、机轮收上着陆、翻转（turnover）、"
                     "客舱质量块约束、机身内发动机约束"],
        ["历史版本数", "7 版（含 1965 年初始版）"],
        ["时间跨度", "1965-02-01 → 2012-01-31（约 47 年）"],
        ["当前状态", "Historical。2017 年 Amdt 23-64 性能化重构后，本条内容由 §23.2270 等性能化条款承接"],
    ], widths=[3.6, 13.2], align_center_cols=(0,))

    doc.add_paragraph()
    H(doc, "1.2　修订沿革总表（按时间排序）", 2)
    mk_table(doc, HIST_HEADERS, HIST_ROWS,
             widths=[0.8, 2.0, 2.0, 3.2, 3.2, 5.6], font=8.5, align_center_cols=(0, 1, 2))

    doc.add_paragraph()
    H(doc, "1.3　修订脉络一览", 2)
    RICH(doc, [
        ("1965 ", 'bold'), ("初版（CAR 3.386 重编）　→　", 'normal'),
        ("1969 ", 'bold'), ("澄清性更正　→　", 'normal'),
        ("1987 ", 'bold'), ("通勤类引入（条文未动）　→　", 'normal'),
        ("1988 ", 'bold'), ("坠撞动力学重构（", 'normal'), ("质量块 18.0g + 翻转客观化", 'rev'), ("）　→　", 'normal'),
        ("1994 ", 'bold'), ("（", 'normal'), ("向下 6.0g", 'new'), ("）　→　", 'normal'),
        ("1996 ", 'bold'), ("（JAR 协调 + ", 'normal'), ("新增 (e)", 'new'), ("）　→　", 'normal'),
        ("2012 ", 'bold'), ("（", 'normal'), ("机身内发动机 18.0g", 'new'), ("）", 'normal'),
    ], size=10.5)

    doc.add_page_break()

    # ================= 2. 深度剖析 =================
    H(doc, "2　各阶段修订背景与原因深度剖析", 1)

    # 2.1
    H(doc, "2.1　初始颁行（1964–1965，Amdt 23-0）：从 CAR 3.386 到 FAR §23.561", 2)

    BULLET(doc, "修订背景：", [
        ("1961 年 11 月 15 日，FAA 以 ", 'normal'),
        ("Draft Release 61-25", 'bold'),
        ("（26 FR 10698）宣布启动民用航空法规重编计划（Recodification Project），把 Civil Air Regulations"
         "（CAR）逐部重编为 Federal Aviation Regulations（FAR）[1][2]。", 'normal'),
    ])
    BULLET(doc, "直接起因：", [
        ("既有 CAR Part 3 由不同年代、不同起草人写成，文体与措辞极不统一，条款编号也无法与 Part 25/27/29 对应。"
         "FAA 明确本次重编的目的", 'normal'),
        ("“仅是梳理与澄清现行规章语言、删除过时或冗余条文”", 'bold'),
        ("，并且", 'normal'),
        ("“未作任何增加公众负担的实质性改动”", 'bold'),
        ("（No substantive changes involving an increased burden on the public）[1]。", 'normal'),
    ])
    BULLET(doc, "本次重编的三条编辑原则：", [
        ("① 统一用 ", 'normal'), ("must", 'bold'), (" 取代 ", 'normal'), ("shall", 'bold'),
        ("——因为适航标准只是颁发型号合格证的前提条件，用命令式的 shall 并不恰当；不符合标准的结果是"
         "不予颁发型号合格证，而非处罚 [2]。", 'normal'),
    ])
    P(doc, "② 重排条款编号，使 Part 23 与 Part 25/27/29 中相对应要求的编号一致；并把散落各处的同类要求"
           "（如飞行手册数据记录、颤振要求）归并到逻辑位置 [2]。", size=10.5, indent=1.1)
    P(doc, "③ 把 CAR 中具备规章性质、但以 CAM（Civil Aeronautics Manual）形式存在的材料一并纳入规章正文；"
           "纯咨询性质的材料转入 Advisory Circular 体系 [2]。", size=10.5, indent=1.1)

    BULLET(doc, "§23.561 的直接来源：", [
        ("在 NPRM 64-17 中，该要求的编号为 ", 'normal'), ("Sec. 23.709 “Protection”", 'bold'),
        ("，条文末尾的修订注释明确写着 ", 'normal'),
        ("“[Revision note: Based on Sec. 3.386]”", 'bold'),
        ("——即源自 CAR 3.386 [1]。在 1964 年 12 月 18 日的最终规章（29 FR 17955）中，该条被重新编号为 ", 'normal'),
        ("Sec. 23.561 “General.”", 'bold'), (" [2]。", 'normal'),
    ])

    BULLET(doc, "NPRM → Final Rule 的两处可核查差异：", [
        ("① 载荷表 ", 'normal'), ("Upward", 'bold'), (" 由 NPRM 中的 ", 'normal'), ("2.0g", 'del'),
        (" 改为最终规章的 ", 'normal'), ("3.0g", 'rev'), ("（正常类/实用类；特技类两版均为 4.5g）[1][2]。", 'normal'),
    ])
    P(doc, "② (e) 段的豁免引用由 NPRM 的 Sec. 23.727 改为最终规章的 Sec. 23.787 [1][2]。",
      size=10.5, indent=1.1)

    BULLET(doc, "技术内涵（初始版五段结构）：", [
        ("(a) 总体要求——飞机虽可能在应急着陆中受损，仍须按本节设计以保护每位乘员；(b) 以 ", 'normal'),
        ("“minor crash landing”（轻微坠撞着陆）", 'bold'),
        (" 为设计场景，给出乘员极限静惯性力表：Upward 3.0g / 4.5g（特技类）、Forward 9.0g、Sideward 1.5g；"
         "(c) 可收放起落架的机轮收上着陆；(d) 翻转（turnover）；(e) 支撑结构须约束可能伤人的"
         "“items of mass”（质量块）[2]。", 'normal'),
    ])

    # 2.2
    H(doc, "2.2　Amdt 23-7（1969）：两处澄清性更正", 2)
    BULLET(doc, "修订背景：", [
        ("1965 年 10 月 25–29 日，当时的 Federal Aviation Agency 召开了 ", 'normal'),
        ("Agency-Industry 会议", 'bold'),
        ("，审查与小飞机制造有关的规章；Notice 67-14（32 FR 5791）正是这次会议的产物之一 [3]。", 'normal'),
    ])
    BULLET(doc, "修订原因（NPRM 67-14 Proposal 17 原文）：", [
        ("变更 1：", 'bold'),
        ("在 (c)(3) 与 (d) 引导句的 “assuming” 之后加入 ", 'normal'),
        ("“, in the absence of a more rational analysis —”", 'new'),
        ("。理由是", 'normal'),
        ("“反映当前立法意图：允许使用经合理分析确定的极限惯性力与摩擦系数，而不必拘泥于条款给定的数值”", 'bold'),
        (" [3]。", 'normal'),
    ])
    BULLET(doc, "", [
        ("变更 2：把 (c)(3)(i) 的 ", 'normal'), ("“upward”", 'del'), (" 删除并替换为 ", 'normal'),
        ("“downward”", 'rev'),
        ("。NPRM 给出的理由是：", 'normal'),
        ("CAR Sec. 3.386(c) 后注释中的 “vertical ultimate acceleration”，本意是指在机轮收上的应急着陆状态中、"
         "相对于飞机向下（downward）的惯性力；FAR 沿用时误写为 upward", 'bold'),
        (" [3]。", 'normal'),
    ])
    BULLET(doc, "最终处置：", [
        ("34 FR 13078（1969-08-13）按提案原样采纳这两处修改 [4]。二者均为澄清性与纠偏性修改，"
         "不改变任何数值，不增加符合性负担；但它们为后续“允许用更合理的分析替代条款给定值”打开了制度口子，"
         "成为 1988 年翻转判据客观化的先声。", 'normal'),
    ])

    # 2.3
    H(doc, "2.3　Amdt 23-34（1987）：通勤类引入与“未改之改”", 2)
    BULLET(doc, "修订背景：", [
        ("自 1953 年起，适航标准以 12,500 磅最大取证起飞重量划分小飞机与大飞机，不考虑运行种类；"
         "当这一分界确立时，人们并未预料到它会成为通勤运行的瓶颈 [5]。1966 年 FAA 设立 ", 'normal'),
        ("air taxi 适航计划", 'bold'),
        ("，用 special conditions 与 SFAR 41 / 41A / 41B / 41C 系列作为从 Part 23 向 Part 25 过渡的临时安排 [9]。"
         "Notice 83-17（48 FR 52010）据此提出正式设立", 'normal'),
        ("通勤类（Commuter Category）", 'bold'), (" [5]。", 'normal'),
    ])
    BULLET(doc, "本版对 §23.561 的实际处理——三份证据：", [
        ("① NPRM 83-17 未提出对 §23.561 的任何修改 [5]。", 'normal'),
    ])
    P(doc, "② 评论环节有评论者指出：通勤类飞机未提出强化 §23.561 应急着陆载荷因子，并援引 "
           "NTSB 1981 年报告《Cabin Safety in Transport Category Aircraft》，认为 §23.561 的极限惯性力"
           "“远低于人体能够承受的水平”。FAA 答复：本次提案的目的不是在此刻重新评估 Part 23 所有飞机的"
           "应急着陆载荷因子；FAA 将在 Part 23 Airworthiness Review 框架下考虑修订 §23.561；"
           "该意见超出 Notice 83-17 的范围，本次不予处理 [6]。", size=10.5, indent=1.1)
    P(doc, "③ 最终规章 52 FR 1806（1987-01-15）共 43 条修订指令，涉及 23.1、23.3、23.25、23.45、…、23.1587 等条款，"
           "其中不含 §23.561 [6]。", size=10.5, indent=1.1)

    info_box(doc, "结论与提示", [
        "Amdt 23-34 的最终规章本身未改动 §23.561 的条文。DRS 之所以把 §23.561 的一个版本归到 "
        "Amdt 23-34，是因为该修订新增的通勤类条款（§23.785(g)、§23.807(d)、§23.963(f) 等）"
        "开始交叉引用 §23.561。",
        "本版对 §23.561 的实际意义是：适用范围正式扩展至通勤类飞机。",
        "更重要的是，上面那条“载荷因子过低”的评论被 FAA 明确记入案卷，直接成为 1986 年 Notice 86-19 "
        "（小飞机适航审查计划第一号通告）的伏笔——见 2.4。",
        "⚠ 但同年 9 月 14 日的更正通告 52 FR 34745 确实改到了 §23.561 —— 见 2.8。",
    ], bg=C_SUB_BG)

    # 2.3.1 修订案号分歧考据
    H(doc, "2.3.1　考据：CFR 溯源注为何把 52 FR 34745 标成 Amdt 23-24", 3)
    P(doc, "§23.561 在 CFR 官方溯源注里写作“Amdt. 23-24, 52 FR 34745, Sept. 14, 1987”，"
           "但 FAA DRS 把 §23.561 的 1987 年版归在 Amdt 23-34（生效 1987-02-17），"
           "而 DRS 的 Amdt 23-24 生效日为 1979-12-31。两个源对不上 [16][17]。", size=10.5)
    BULLET(doc, "分歧根因（已核实到原文）：", [
        ("52 FR 34745 的 ", 'normal'), ("FR 抬头", 'bold'),
        ("把修订案号印成 “[Docket No. 23516; Amdt. Nos. 21-59, ", 'normal'),
        ("23-24", 'del'), (", 36-13, 91-197, and 135-21]”；而同一份文件的 ", 'normal'),
        ("SUPPLEMENTARY INFORMATION 与 “Correction of Publication” 两处均写 23-34", 'bold'),
        ("，且开篇即说明是更正 1987-01-15 刊登于 52 FR 1806 的最终规章 [17]。", 'normal'),
        ("1987-03-09 的同一案卷更正件 52 FR 7261 抬头正确印作 ", 'normal'), ("23-34", 'bold'),
        ("，可交叉印证 [17]。", 'normal'),
        ("→ 结论：这是 FR 印刷错误，CFR 编纂时照抄了错误的抬头，遂把错误带进官方溯源注。"
         "应以 DRS 的 Amdt 23-34 为准。", 'normal'),
    ])

    # 2.4
    H(doc, "2.4　Amdt 23-36（1988）：坠撞动力学与静态/动态分工", 2)
    BULLET(doc, "修订背景（完整链条）：", [
        ("1980 年代中期，FAA 启动 ", 'normal'), ("Small Airplane Airworthiness Review Program", 'bold'),
        ("，并召开小飞机适航审查会议（Small Airplane Airworthiness Review Conference）；"
         "通用航空安全小组（GASP）在该会议上提出第 218、219、220、221、222、518 号提案 [7][8]。"
         "Notice 86-19（51 FR 44878, 1986-12-12）据此提出修改 §23.561、23.783、23.785、23.787、23.807，"
         "并新增 §23.562 动态着陆条件 [7]。", 'normal'),
    ])
    BULLET(doc, "技术动因：", [
        ("FAA 在 NPRM 中直言：“", 'normal'),
        ("Section 23.561 is proposed for revision since proposed Section 23.562 proposes the use of "
         "dynamic requirements and criteria for occupant protection. Dynamic requirements more "
         "accurately simulate emergency landing conditions and the timeframe in which the protection "
         "of occupants is intended.", 'bold'),
        ("”——即静态载荷因子无法反映真实事故中随时间变化的动态冲击响应 [7]。", 'normal'),
    ])

    H(doc, "2.4.1　(b) 段全面重写", 3)
    BULLET(doc, "① 术语统一（评论驱动）：", [
        ("有评论者建议把 (b) 中的 ", 'normal'), ("“minor crash landing”", 'del'), (" 换成 ", 'normal'),
        ("“emergency landing conditions”", 'rev'), ("，并把 ", 'normal'), ("“escape serious injury”", 'del'),
        (" 换成 ", 'normal'), ("“designed to protect each occupant”", 'rev'),
        ("。FAA 同意并采用该措辞 [8]。", 'normal'),
    ])
    BULLET(doc, "② 载荷表按类别重写：", [
        ("因 Amdt 23-34 已引入通勤类，(b)(2)(i) 明确写为 ", 'normal'),
        ("“Upward, 3.0g for normal, utility, and commuter category airplanes, or 4.5g for acrobatic "
         "category airplanes”", 'rev'),
        ("；并把“一般适用于所有 Part 23 飞机类别”的应急着陆极限静载荷因子集中放入 (b)(2) [8]。", 'normal'),
    ])
    BULLET(doc, "③ 新增 (b)(3) 客舱质量块载荷：", [
        ("Upward 3.0g / ", 'new'), ("Forward 18.0g", 'new'), (" / Sideward 4.5g", 'new'), ("。", 'normal'),
    ])
    P(doc, "技术依据：FAA 与 NASA 联合开展的全尺寸坠撞试验表明，可生存坠撞中纵向冲击载荷可超过 9g"
           "（NASA Technical Paper 2083，1982 年 11 月，《Correlation and Assessment of Structural "
           "Airplane Crash Data With Flight Parameters at Impact》）。GASP 原先建议对质量块做动态试验，"
           "但 FAA 与 GASP 随后评估认为：按建议的极限静载荷做静力试验即可实现原动态试验的意图，"
           "且能显著降低申请人的取证成本 [8]。", size=10.5, indent=1.1)
    BULLET(doc, "评论处置（质量块载荷）：", [
        ("· 有评论者认为既有 9g 前向要求已足够、反对 18g。FAA 不同意：9g 不足以达到 Notice 86-19 "
         "提出的客舱安全水平 [8]。", 'normal'),
        ("· 有评论者建议增加 1.5g 后向（rebound）载荷，以防后向隔舱等内容物回落砸伤下方乘员。"
         "FAA 表示理解并将在未来规则制定中评估，本次不采纳 [8]。", 'normal'),
        ("· 有评论者建议 Upward 改 4.5g、Sideward 改 4.5g 并加后向 1.5g。FAA 以“无技术依据且超出 "
         "notice 范围”不予采纳 [8]。", 'normal'),
        ("· 有评论者建议改为“把质量块布置在松脱后不会伤人/不会穿透油路/不会堵住出口的位置”。"
         "FAA 不同意：对许多小飞机设计而言这不现实；更高的载荷因子是防止坠撞中质量块松脱伤人的必要手段 [8]。", 'normal'),
    ])

    H(doc, "2.4.2　与新增 §23.562 的分工", 3)
    P(doc, "本次修订的技术支点，是把动态试验（26g/19g 试验条件、170 磅 ATD、HIC 头部损伤判据等）交给新增的 "
           "§23.562，§23.561 保留为静态/通用要求，两者互补。其技术支撑包括：NASA 全尺寸坠撞“跌落试验”系列、"
           "NTSB 事故评估信息、FAA 民用航空医学研究所（CAMI）的动态试验，以及军方与汽车领域通用的人体冲击"
           "损伤判据 [8]。", size=10.5, indent=0.3)
    BULLET(doc, "关键边界（FAA 明确拒绝的一项建议）：", [
        ("有评论者要求把机身结构一并纳入动态试验、并测量座椅/约束系统与机体的连接载荷。FAA 拒绝："
         "坠撞时机身地板结构会在严重局部载荷下屈服，从而改变连接界面的载荷；把机身结构纳入动态试验会"
         "大幅增加复杂性与成本；且机身结构本就必须满足 §23.561 的静强度要求。因此最终规章不要求测量"
         "连接载荷（但测量本身可作为有用的设计信息）[8]。", 'normal'),
    ])

    H(doc, "2.4.3　(d) 段翻转判据客观化", 3)
    BULLET(doc, "要解决的歧义：", [
        ("原 (d) 以 ", 'normal'), ("“If a turnover is reasonably probable”", 'del'),
        (" 表述，而 “reasonably probable”（相当可能）在本节中从未定义，长期是条文的歧义来源 [7]。", 'normal'),
    ])
    BULLET(doc, "最终条文：", [
        ("改为 ", 'normal'),
        ("“If it is not established that a turnover is unlikely during an emergency landing…”", 'rev'),
        ("，并给出 (d)(1) 五项分析条件——(i) 最大重量、(ii) 最前重心、(iii) 纵向 9.0g、(iv) 垂直 1.0g、"
         "(v) 前三轮起落架飞机前起支柱失效且机头触地；以及 (d)(2) 翻转后倒置机体的载荷——向上 3.0g、"
         "地面摩擦系数 0.5 [8]。", 'normal'),
    ])
    BULLET(doc, "评论处置（翻转判据）：", [
        ("· 有评论者认为按此判据几乎所有前三轮起落架飞机都会被判定为“可能翻转”。FAA 答复：虽然前三轮"
         "起落架飞机在正常运行中天然抗翻转，但 FAA 事故/事件数据显示，前三轮起落架小飞机在应急着陆"
         "（含目测过低、目测过高、失去方向控制等情形）中确有翻转；高翼前三轮起落架飞机的翻转频率高于"
         "低翼，但低翼飞机的翻转频率同样高到需要评估。此外，一些新型低翼前三轮起落架小飞机采用滑动式"
         "透明座舱盖，在翻转中几乎不提供乘员保护 [8]。", 'normal'),
        ("· 有评论者要求把判据写成“最不利的全机重量与重心组合”。FAA 部分拒绝：all-up-weight 不是 Part 23 "
         "中的定义术语；最前重心位置通常对翻转评估最临界 [8]。", 'normal'),
        ("· 有评论者建议把提案中的 (c)(6) 单独列为 (d)。FAA 同意：(1)–(5) 是判定翻转可能性的判据，"
         "而 (6) 定义的是翻转后施加于倒置机体的载荷，性质不同。据此对 (c)/(d) 重新编号并重新排版 [8]。", 'normal'),
        ("· FAA 还澄清：翻转保护应按静态条件评估，不按 §23.562 那样的动态条件评估 [8]。", 'normal'),
    ])

    H(doc, "2.4.4　(e) 段的去向", 3)
    P(doc, "NPRM 86-19 曾提议删除 (c)（机轮收上着陆），理由是新 §23.562 的动态要求“显著超出”既有要求，"
           "能为可收放起落架飞机的“机腹着陆”提供更现实的乘员保护标准；并把 (d)、(e) 重新编号。最终规章"
           "保留了 (c)，并把质量块的载荷要求并入新的 (b)(3)。需要说明的是，1996 年 Amdt 23-48 的修订指令"
           "措辞为“adding a new paragraph (e)”（新增 (e) 段），这说明在 1988–1996 年间的 CFR 文本中，"
           "§23.561 并不存在独立的 (e) 段 [7][8][12]。", size=10.5, indent=0.3)

    # 2.5
    H(doc, "2.5　Amdt 23-46（1994）：替代应急出口与向下 6.0g", 2)
    BULLET(doc, "修订背景（产业豁免压力）：", [
        ("Amdt 23-34 在 §23.807(d)(1) 要求通勤类飞机（15 座及以下）除旅客登机门外，在客舱每侧各设一个"
         "应急出口。此后", 'normal'),
        ("多家制造厂与改装商申请豁免", 'bold'),
        ("，希望改用 §25.807(c)(1) 的运输类出口构型，并主张通勤类飞机已有其他补偿性客舱安全设计"
         "（更大的出口、更宽的过道等）[9][10]。", 'normal'),
    ])
    BULLET(doc, "FAA 的对比审查：", [
        ("FAA 对 Part 23 通勤类与 Part 25 小型运输类的客舱安全标准做了系统对比，发现 Part 25 的替代出口"
         "构型配有应急照明、更宽过道、额外出口标志等一揽子补偿措施。据此 Notice 90-20（55 FR 35544）提出"
         "“替代应急出口”（§23.807(d)(4)）方案，并配套提出 §23.783(f)、§23.803(b)、§23.811(c) 与新增"
         "§23.812 应急照明等要求 [9]。", 'normal'),
    ])
    BULLET(doc, "为什么把向下载荷移进 §23.561（NPRM 原文）：", [
        ("NPRM 90-20 明确指出：", 'normal'),
        ("§23.561(b)(2) 当时定义的应急着陆条件中并不包含通勤类飞机的向下静惯性载荷", 'bold'),
        ("；此前向下载荷因子是作为“豁免 §23.807(d)(1) 出口数量要求”的附加适航要求，规定在 §23.807(d)(1) 中。"
         "为与运输类标准统一、并简化替代出口标准，FAA 提议把它移入 §23.561(b)(2) 作为新的 (iv) [9]。", 'normal'),
    ])
    BULLET(doc, "最终条文与立法目的：", [
        ("新增 ", 'normal'),
        ("“(iv) Downward, 6.0g when certification to the emergency exit provisions of Sec. 23.807(d)(4) "
         "is requested; and”", 'new'),
        ("。FAA 在 Proposal 2 中说明其目的是", 'normal'),
        ("“确保特定的最小向下机体强度，以保护乘员免于因结构失效而无法撤离”", 'bold'), (" [10]。", 'normal'),
    ])
    BULLET(doc, "评论处置：", [
        ("· 有评论者认为 6g 对 Part 23 飞机过大。FAA 不同意，并援引 1970 年代 FAA 与 NASA 的大量小飞机"
         "坠撞动力学研究——NASA 对单发与双发飞机做过 21 次受控全尺寸冲击试验 [10]。", 'normal'),
        ("· 有评论者建议采纳 NPRM 中作为替代方法提出的更高向下惯性载荷与更高下降率。FAA 不同意："
         "虽然应急着陆时向下载荷因子可能大于 6g，但 FAA 的意图是", 'normal'),
        ("使通勤类飞机的适航标准与小型运输类飞机保持一致", 'bold'),
        ("，故 6g 保留为最低要求。相应地，最终规章", 'normal'),
        ("删除了提案中“任何更小的力”的表述", 'rev'),
        ("，使 Part 23 通勤类的替代客舱安全与应急出口要求与 Part 25 小型运输类进一步统一"
         "（Part 25 中同样删除了该替代力）[10]。", 'normal'),
        ("· 共 5 位评论者对 Notice 90-20 作出回应，FAA 依评论另作了少量技术与编辑性修改 [10]。", 'normal'),
    ])

    # 2.6
    H(doc, "2.6　Amdt 23-48（1996）：FAR/JAR 协调统一", 2)
    BULLET(doc, "修订背景（协调机制时间线）：", [
        ("1990 年 6 月，JAA 理事会与 FAA 会议上，FAA 局长承诺支持 FAR 与欧洲 JAA 正在制定的 JAR 协调统一；"
         "FAA 小飞机局随即设立 ", 'normal'), ("FAA Harmonization Task Force", 'bold'), (" [11][12]。", 'normal'),
    ])
    P(doc, "1990 年 10 月，FAA 协调工作组与 GAMA 委员会在布鲁塞尔参加 JAR 23 研究组会议，欧洲机体制造商"
           "协会 AECMA 亦派代表参加；此后 GAMA、AECMA、FAA、JAA 四方技术代表多次会晤，逐条消除 JAR 与 "
           "Part 23 的差异 [11]。", size=10.5, indent=1.1)
    P(doc, "1991 年 1 月 FAA 设立航空规则制定咨询委员会（ARAC，56 FR 2190）；1992 年 6 月多伦多 "
           "JAA/FAA 协调会议上宣布把协调工作并入 ARAC 架构，由 GABA（通用航空与公务机）议题下的 ARAC 承担；"
           "1992 年 11 月 30 日公布 JAR/FAR 23 协调工作组成立公告（57 FR 56626），1993 年 2 月 2 日首次会议 [11]。",
      size=10.5, indent=1.1)
    P(doc, "因修订量过大，FAA 将 Part 23 协调拆成四个 NPRM（系统与设备、动力装置、飞行、机体），"
           "Notice 94-20（59 FR 35196, 1994-07-08）为“机体”部分 [11]。", size=10.5, indent=1.1)

    BULLET(doc, "§23.561 的协调目标（Notice 94-20 原文）：", [
        ("“修改 (b)、(d) 并新增 (e)，以与 JAR 23 协调。(b) 关于乘员保护，采用与适用于大型飞机的 "
         "Part 25 / JAR 25 相似的语言；(d) 关于翻转，在不做实质改变的前提下简化与澄清；"
         "新增 (e) 关于支撑结构，确保可能伤人的质量块被约束。”[11]", 'bold'),
    ])
    BULLET(doc, "最终落地的三处修改（61 FR 5138）：", [
        ("① (b) 引导句改为 ", 'normal'),
        ("“The structure must be designed to give each occupant every reasonable chance of escaping "
         "serious injury when—”", 'rev'),
        ("，与 Part 25 / JAR 25 一致（1988 版该句为“…to protect each occupant during emergency landing "
         "conditions when:”）[12]。", 'normal'),
        ("② (d)(1) 由五项精简为四项：(i) 最不利的重量与重心组合（合并了 1988 版的“最大重量”与"
         "“最前重心”）、(ii) 纵向 9.0g、(iii) 垂直 1.0g、(iv) 前三轮起落架飞机前起支柱失效且机头触地 [12]。", 'normal'),
        ("③ 新增 (e) 支撑结构约束质量块 [12]。", 'normal'),
    ])
    info_box(doc, "1996 版 (e) 与 1965 初始版 (e) 的对照", [
        [("1965 版：", 'bold'),
         ("“Except as provided in Sec. 23.787 the supporting structure must be designed to restrain, "
          "under loads up to those specified in paragraph (b)(2) of this section, …”", 'normal')],
        [("1996 版：", 'bold'),
         ("“Except as provided in Sec. 23.787(c), the supporting structure must be designed to restrain, "
          "under loads up to those specified in paragraph (b)(3) of this section, …”", 'new')],
        "差异仅两处：豁免引用由 §23.787 精确到 §23.787(c)；载荷引用由 (b)(2) 更新为 (b)(3)"
        "——后者正是 1988 年把质量块载荷单列为 (b)(3) 后必然的连带更正。",
    ], bg=C_SUB_BG)

    # 2.7
    H(doc, "2.7　Amdt 23-62（2012）：机身内嵌发动机与性能化转向", 2)
    BULLET(doc, "修订背景（ARC 路径）：", [
        ("2003 年 2 月 3 日，FAA 公布成立 ", 'normal'), ("part 125/135 Aviation Rulemaking Committee（ARC）", 'bold'),
        ("（68 FR 5488）；ARC 于 2005 年完成工作，就 Part 23 涡轮喷气飞机的安全标准提出建议，"
         "涉及 41 个 Part 23 条款，并复查了从轻型活塞多发飞机到小型喷气机的事故史。"
         "FAA 的目标是把此前通过 special conditions 逐案施加的要求法典化 [13][14]。", 'normal'),
    ])
    BULLET(doc, "§23.561 的具体起因（NPRM 原文）：", [
        ("“FAA 提议把近期适用于某型单发涡轮喷气机的 turbojet special condition 法典化写入 §23.561。"
         "本提案适用于机身内嵌中置发动机的单发涡轮喷气机。Part 23 原本没有覆盖机身内嵌中置发动机的安装形式，"
         "唯一的例外是同轴推进式螺旋桨构型。鉴于若干新喷气机设计的出现，对位于客舱后方的机身内发动机"
         "提出更高的发动机保持强度要求是审慎的。”[13]", 'bold'),
        ("立法目的：", 'normal'),
        ("降低发动机在前向坠撞载荷下从安装架脱落、进而侵入客舱的可能性", 'bold'), (" [13][14]。", 'normal'),
    ])
    BULLET(doc, "最终条文：", [
        ("(e)(1) ", 'new'),
        ("“For engines mounted inside the fuselage, aft of the cabin, it must be shown by test or analysis "
         "that the engine and attached accessories, and the engine mounting structure—"
         "(i) Can withstand a forward acting static ultimate inertia load factor of 18.0 g plus the "
         "maximum takeoff engine thrust; or "
         "(ii) The airplane structure is designed to preclude the engine and its attached accessories "
         "from entering or protruding into the cabin should the engine mounts fail.”", 'new'),
        ("　(e)(2) [Reserved] [14]", 'new'),
    ])

    H(doc, "2.7.1　NPRM → Final Rule 的四处评论驱动变化（本节最有价值的部分）", 3)
    mk_table(doc, ["提出方", "NPRM 提案 / 原表述", "FAA 处置", "最终条文变化"], [
        ["EASA",
         "“For turbojet engines mounted inside the fuselage, aft of the cabin…”",
         "采纳。理由：任何以该构型安装的发动机在应急着陆时都可能危及客舱乘员，规章不应限于涡轮喷气发动机。",
         [("“turbojet engines” → “engines”", 'rev')]],
        ["Transport Canada",
         "建议在飞机 VS0 超过 61 节时上调载荷因子",
         "拒绝。理由：本条要求发动机在 18 g 加上最大起飞推力的组合下被保持，对发动机保持而言是合理的。",
         "维持 18.0 g"],
        ["Transport Canada",
         "建议附件（accessories）不必承受最大起飞推力的附加载荷",
         "拒绝。理由：虽然附件不直接反作用推力载荷，但附件会对其安装结构施加载荷，且通常在发动机大功率输出时最大。",
         "维持 “18.0 g plus the maximum takeoff engine thrust”"],
        ["Transport Canada",
         "“(ii) to deflect the engine … away from the cabin” 表述过于局限",
         ("采纳。理由：设计者可能提出其他方法（如吸能隔框/屏障）；采纳后条款更性能化，"
          "可避免“指定机体设计”（dictation of the airframe design）。"),
         [("改为 “preclude the engine … from entering or protruding into the cabin”", 'rev')]],
    ], widths=[2.2, 4.6, 5.4, 4.6], font=8.5)

    doc.add_paragraph()
    BULLET(doc, "同期配套修订：", [
        ("Amdt 23-62 还修订了 §23.562（要求通勤类喷气机进行座椅动态试验、HIC 计算与 §25.562 一致），"
         "并把通勤类扩展到涡轮喷气飞机。FAA 明确说明：本次把 §23.562 适用于所有通勤类喷气机，"
         "是“采纳现行产业实践”（所有新取证的通勤类喷气机制造商都已同意遵守 §23.562），"
         "并刻意不把通勤类螺旋桨飞机纳入 [14]。", 'normal'),
    ])

    # 2.8 更正通告与 CFR 编纂校正
    H(doc, "2.8　三份被溯源注遗漏的“更正”类文献（1987 / 2007）", 2)
    P(doc, "CFR 官方溯源注只登记修订案的 FR 引证，更正通告往往被吞掉或只留一个裸页码，"
           "CFR 编纂校正（OFR 在年度版排印时发的 Correction）更是完全不进溯源注。"
           "§23.561 恰好三种情况都有，且都对条文产生了实际影响 [17][18][19]。", size=10.5)

    mk_table(doc, ["FR 引证", "日期", "性质", "对 §23.561 的实际改动"], [
        ["52 FR 34745", "1987-09-14", "FAA 更正通告",
         "新增修订指令 25-1：§23.561(b)(2) 表格第一栏标题 "
         "“Normal and utility categories” → “Normal, utility, and commuter categories” [17]"],
        ["72 FR 59939", "2007-10-23", "CFR 编纂校正",
         "CFR 2007 版第 227 页 §23.561(d)(1) 重复排印了一遍 (i)–(iv)，"
         "并多出 (d)(1)(v)；OFR 指令删除重复的第二组及 (d)(1)(v) [18]"],
        ["72 FR 72915", "2007-12-26", "CFR 编纂校正（重发）",
         "上一件表述不够精确，OFR 重新发布：删除自第二个 (d)(1)(i) 起的连续五段 [19]"],
    ], widths=[2.6, 2.2, 3.2, 8.8], font=8.5)

    BULLET(doc, "为什么这三份值得单独列出：", [
        ("① 52 FR 34745 是 §23.561 迄今唯一一次“因印刷错误而改条文”的记录 —— "
         "它把通勤类补进了载荷表的类别栏，与 Amdt 23-34 引入通勤类是同一件事的收尾 [17]。", 'normal'),
        ("② 两份 2007 年 CFR Correction 表明：CFR 年度版曾把 (d)(1) 整段重排过一次，"
         "造成 (i)–(iv) 与 (v) 重复出现。这不是规章变更，而是 GPO 排印错误；"
         "引用 2007 年印本（含部分电子数据库）时须注意该重复文本已被官方删除 [18][19]。", 'normal'),
        ("③ 这三份都不在 DRS 的条款修订版索引里，也不在 CFR 溯源注的修订案序列里；"
         "只有逐份比对 FR 原文才能发现。", 'normal'),
    ])

    doc.add_page_break()

    # ================= 3. 条文演变对比 =================
    H(doc, "3　条款原文演变对比", 1)
    P(doc, "本节先给出颜色图例，再按版本逐条列出条文（含英文原文），最后以要素矩阵横向对照七个版本。"
           "每张条文表的“状态”列标明该段在本版是新增、修订、删除还是继承未变。", size=10.5)
    legend(doc)

    clause_table(doc, "3.1　版本一：Amdt 23-0（初始版，1965-02-01 生效）　　来源：29 FR 17955 [2]",
                 V1_ROWS,
                 note="说明：NPRM 64-17 中该条编号为 Sec. 23.709 “Protection”，Upward 提案值为 2.0g；"
                      "最终规章改为 3.0g [1][2]。")

    clause_table(doc, "3.2　版本二：Amdt 23-7（1969-09-14 生效）　　来源：34 FR 13078 [4]",
                 V2_ROWS,
                 note="两处修改均为澄清性/纠偏性：加入“合理分析”许可，并把 (c)(3)(i) 的 upward 更正为 downward [3][4]。")

    P(doc, "3.3　版本三：Amdt 23-34（1987-02-17 生效）　　来源：52 FR 1806 [6]",
      size=11, bold=True, color=RGBColor(0x1F, 0x4E, 0x79), before=10, after=3)
    info_box(doc, "本版 §23.561 条文与 Amdt 23-7 版完全一致（未作任何文字修改）", [
        "依据：52 FR 1806 的全部 43 条修订指令中不含 §23.561；NPRM 83-17 亦未提出修改 §23.561 [5][6]。",
        "本版的实际影响：通勤类（Commuter Category）正式进入 Part 23，§23.561 的适用范围随之扩展；"
        "新增的 §23.785(g)、§23.807(d)、§23.963(f) 等通勤类条款开始交叉引用 §23.561 的应急着陆载荷。",
        "同时，FAA 在评论答复中记录了“§23.561 载荷因子远低于人体可承受水平”的意见，"
        "并承诺在 Part 23 Airworthiness Review 框架下处理——这一承诺在一年后的 Notice 86-19 中兑现。",
    ], bg="E2EFDA")

    clause_table(doc, "3.4　版本四：Amdt 23-36（1988-09-14 生效）　　来源：53 FR 30802 [8]",
                 V4_ROWS,
                 note="(b) 全面重写、(d) 客观化；(c) 保留；(e) 因重组而不再独立存在。")

    clause_table(doc, "3.5　版本五：Amdt 23-46（1994-06-16 生效）　　来源：59 FR 25766 [10]",
                 V5_ROWS,
                 note="本版仅新增 (b)(2)(iv) 一项；其余条文与 Amdt 23-36 版一致。")

    clause_table(doc, "3.6　版本六：Amdt 23-48（1996-03-11 生效）　　来源：61 FR 5138 [12]",
                 V6_ROWS,
                 note="FAR/JAR 协调版。(b) 引导句回到 Part 25 表述；(d)(1) 五项精简为四项；新增 (e)。")

    clause_table(doc, "3.7　版本七：Amdt 23-62（2012-01-31 生效）　　来源：76 FR 75736 [14]，"
                      "并以 eCFR 2017-01-01 版校对 [15]",
                 V7_ROWS,
                 note="新增 (e)(1) 机身内发动机坠撞约束与 (e)(2) [Reserved]；其余与 Amdt 23-48 版一致。")

    H(doc, "3.8　关键条款要素演变矩阵（横向对照七个版本）", 2)
    mk_table(doc, MATRIX_HEADERS, MATRIX_ROWS,
             widths=[3.0, 1.9, 1.9, 1.9, 2.1, 2.0, 2.0, 2.0], font=8,
             align_center_cols=tuple(range(1, 8)))

    doc.add_page_break()

    # ================= 4. 总结 =================
    H(doc, "4　总结：§23.561 四十余年的演进规律", 1)
    mk_table(doc, ["驱动力", "在 §23.561 上的体现", "对应版本"], [
        ["事故调查与研究驱动",
         "NTSB《Cabin Safety in Transport Category Aircraft》报告与 FAA/NASA 全尺寸坠撞试验"
         "（NASA TP 2083）直接支撑了质量块 18.0g 与翻转判据客观化", "23-36（1988）"],
        ["产业豁免压力倒逼",
         "多家制造商申请豁免 §23.807(d)(1) 出口数量要求，FAA 以“替代应急出口 + 向下 6.0g”"
         "一揽子方案回应", "23-46（1994）"],
        ["国际协调",
         "JAA/EASA/Transport Canada 的逐条评论改变了条文措辞："
         "(b) 引导句与 Part 25/JAR 25 对齐；发动机条款由 turbojet 扩展到全部发动机；"
         "“deflect the engine”改为性能化的“preclude … from entering or protruding”",
         "23-48（1996）\n23-62（2012）"],
        ["新技术/新构型",
         "通勤类（1987）、通勤类喷气机与机身内嵌中置发动机（2012）依次被纳入",
         "23-34 / 23-62"],
        ["从处方性走向性能化",
         "2012 年 FAA 明确表示采纳 Transport Canada 意见是为了“使规则更性能化、"
         "避免指定机体设计”（preclude dictation of the airframe design）", "23-62（2012）"],
        ["从单一静态走向静态+动态分工",
         "§23.561 保留静态通用要求，新增 §23.562 承担动态试验（26g/19g、ATD、HIC）",
         "23-36（1988）"],
    ], widths=[3.2, 9.6, 4.0], font=9)

    doc.add_paragraph()
    info_box(doc, "两点提示", [
        "① 本报告覆盖 Amdt 23-63 及更早的历史版本，与 DRS 快照「Part 23 @ Amdt 23-63 as of 08/29/2017 "
        "and earlier」一致。2017 年 8 月 30 日生效的 Amdt 23-64 对 Part 23 作了性能化重构，"
        "原 §23.561 的静态 g 值表格被安全目标导向的条款（如 §23.2270 乘员应急防护）取代，"
        "具体数值与符合性方法转入 ASTM 等共识标准。该重构不在本快照范围内，如需可另出一份对照报告。",
        "② 若后续按此模板批量生成 Part 23 全部条款的修订史，建议固定本节的结构："
        "「汇总表 → 逐版背景（含 NPRM 提案原文与 FAA 评论处置）→ 条文彩色对比 → 要素矩阵 → 参考文献」。",
    ], bg=C_NOTE_BG)

    # ================= 5. 参考文献 =================
    H(doc, "5　参考文献", 1)
    P(doc, "下列全部文件均已下载并保存于本地目录「FAR 23/Rulemaking Docs」；"
           "FR 引证格式为「卷号 FR 起始页码」。文中 [n] 编号与本表一致。", size=10)
    mk_table(doc, ["编号", "类型", "文件编号 / 案号", "Federal Register 引证与日期", "本地文件名 / 来源"],
             REF_ROWS, widths=[1.0, 1.6, 4.2, 4.6, 5.4], font=8.5,
             align_center_cols=(0, 1))

    doc.add_paragraph()
    H(doc, "附录　DRS 条款版本记录 ID", 2)
    P(doc, "FAA 动态法规系统（DRS）中 §23.561 各历史版本的记录 ID，"
           "可用于在 DRS 中直接定位相应版本条文。", size=10)
    mk_table(doc, ["修订号", "生效日期", "DRS 记录 ID（docUniqueId）"], DRS_ROWS,
             widths=[3.0, 3.0, 10.8], font=9, align_center_cols=(0, 1))

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("已生成:", OUT)


if __name__ == "__main__":
    build()
