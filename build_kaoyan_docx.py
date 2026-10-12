# -*- coding: utf-8 -*-
"""
生成《2027 考研管综数学（199）真题错题分析与高效备考方案》Word 文档。

依赖：python-docx
    pip install python-docx
运行：
    python build_kaoyan_docx.py
输出：
    脚本同目录下的 2027考研管综数学199错题分析与高效备考方案.docx
"""

import os
import re
import sys

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.exit("缺少依赖 python-docx，请先执行： pip install python-docx")


# --------------------------------------------------------------------------- #
# 全局样式常量
# --------------------------------------------------------------------------- #
NAVY = RGBColor(0x1F, 0x4E, 0x79)        # 标题深蓝 #1F4E79
NAVY_HEX = "1F4E79"
TEXT = RGBColor(0x1A, 0x1A, 0x1A)        # 正文近黑
GREY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD_TEXT = RGBColor(0x9C, 0x65, 0x00)

EA_BODY = "宋体"                          # 中文正文字体
EA_HEAD = "微软雅黑"                       # 中文标题字体
LATIN = "Calibri"                        # 西文字体

BODY_SIZE = 10.5                         # 五号字
TITLE_SIZE = 20
H1_SIZE = 14
H2_SIZE = 12

GOLD_FILL = "FFF7E0"                     # 重点框底色
GOLD_LINE = "D9A404"
BLUE_FILL = "EAF1F8"
ROW_FILL = "F2F6FB"

BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


# --------------------------------------------------------------------------- #
# 基础工具函数
# --------------------------------------------------------------------------- #
def style_run(run, size=BODY_SIZE, bold=False, italic=False, color=None,
              ea=EA_BODY, latin=LATIN):
    """统一设置 run 的字体（含中文 eastAsia）、字号、颜色、粗斜体。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), ea)
    return run


def add_rich(paragraph, text, size=BODY_SIZE, color=TEXT, bold=False,
             ea=EA_BODY, latin=LATIN):
    """向段落写入文本，支持 **加粗** 内联标记。"""
    pos = 0
    for match in BOLD_RE.finditer(text):
        if match.start() > pos:
            style_run(paragraph.add_run(text[pos:match.start()]),
                      size, bold, False, color, ea, latin)
        style_run(paragraph.add_run(match.group(1)),
                  size, True, False, color, ea, latin)
        pos = match.end()
    if pos < len(text):
        style_run(paragraph.add_run(text[pos:]),
                  size, bold, False, color, ea, latin)
    return paragraph


def fmt_para(paragraph, space_before=0, space_after=4, line=1.25, align=None):
    """统一段落间距与行距（正文 1.25 倍行距）。"""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if align is not None:
        paragraph.alignment = align
    return paragraph


def safe_style(doc, name):
    """样式存在则返回样式名，否则返回 None。"""
    try:
        doc.styles[name]
        return name
    except KeyError:
        return None


def shade_cell(cell, fill):
    """设置单元格底纹颜色。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=90, bottom=90, left=140, right=140):
    """设置单元格内边距（单位 dxa，1/20 磅）。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for tag, value in (("top", top), ("left", left),
                       ("bottom", bottom), ("right", right)):
        node = OxmlElement("w:" + tag)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tc_pr.append(mar)


def set_table_borders(table, color=NAVY_HEX, size=6):
    """为表格添加四边与内部网格线。"""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)
        borders.append(node)
    tbl_pr.append(borders)


def add_hrule(doc, color=NAVY_HEX, size=12):
    """标题下方的深蓝色分隔横线。"""
    paragraph = doc.add_paragraph()
    fmt_para(paragraph, space_before=0, space_after=10, line=1.0)
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    ppr.append(pbdr)
    return paragraph


def add_spacer(doc, pt=6):
    paragraph = doc.add_paragraph()
    fmt_para(paragraph, space_after=pt, line=1.0)
    style_run(paragraph.add_run(""), size=1)
    return paragraph


# --------------------------------------------------------------------------- #
# 文档构件
# --------------------------------------------------------------------------- #
def add_title(doc, title, subtitle):
    p = doc.add_paragraph()
    fmt_para(p, space_before=0, space_after=6, line=1.15,
             align=WD_ALIGN_PARAGRAPH.CENTER)
    style_run(p.add_run(title), size=TITLE_SIZE, bold=True,
              color=NAVY, ea=EA_HEAD)

    ps = doc.add_paragraph()
    fmt_para(ps, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_rich(ps, subtitle, size=BODY_SIZE, color=GREY, ea=EA_HEAD)
    add_hrule(doc)


def add_heading(doc, text, level=1):
    """level=1 → 14pt 加粗；level=2 → 12pt 加粗；均使用深蓝色。"""
    style = safe_style(doc, "Heading 1" if level == 1 else "Heading 2") or "Normal"
    size = H1_SIZE if level == 1 else H2_SIZE
    p = doc.add_paragraph(style=style)
    fmt_para(p, space_before=(14 if level == 1 else 10),
             space_after=(6 if level == 1 else 4))
    add_rich(p, text, size=size, color=NAVY, bold=True, ea=EA_HEAD)
    return p


def add_body(doc, text, space_before=0, space_after=4):
    p = doc.add_paragraph()
    fmt_para(p, space_before=space_before, space_after=space_after)
    add_rich(p, text)
    return p


def add_label(doc, text):
    """小标题式标签（如“涉及题型：”），深蓝色加粗。"""
    p = doc.add_paragraph()
    fmt_para(p, space_before=6, space_after=2)
    add_rich(p, text, color=NAVY, bold=True)
    return p


def add_bullet(doc, text, level=1, space_after=3):
    """项目符号列表。level=1 为一级圆点，level=2 为二级短横。"""
    style = safe_style(doc, "List Bullet" if level == 1 else "List Bullet 2") \
        or safe_style(doc, "List Bullet")
    p = doc.add_paragraph(style=style or "Normal")
    fmt_para(p, space_after=space_after)
    if style is None:
        p.paragraph_format.left_indent = Cm(0.75 * level)
        text = ("• " if level == 1 else "– ") + text
    add_rich(p, text)
    return p


def add_numbered(doc, index, text, space_after=3):
    """编号列表，优先使用 Word 原生 List Number 样式。"""
    style = safe_style(doc, "List Number")
    p = doc.add_paragraph(style=style or "Normal")
    fmt_para(p, space_after=space_after)
    if style is None:
        p.paragraph_format.left_indent = Cm(0.75)
        text = "%d. %s" % (index, text)
    add_rich(p, text)
    return p


def add_callout(doc, title, lines, fill=GOLD_FILL, line_color=GOLD_LINE,
                title_color=GOLD_TEXT):
    """带底色与边框的重点框（用于核心策略、舍弃原则等）。"""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(16.6)
    shade_cell(cell, fill)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    set_table_borders(table, color=line_color, size=12)

    p0 = cell.paragraphs[0]
    fmt_para(p0, space_before=0, space_after=3)
    add_rich(p0, "◆ " + title, size=11, bold=True,
             color=title_color, ea=EA_HEAD)
    for text in lines:
        p = cell.add_paragraph()
        fmt_para(p, space_after=2)
        add_rich(p, text)
    add_spacer(doc, pt=4)
    return table


def add_table(doc, headers, rows, widths):
    """构建深蓝表头的规整表格；widths 为各列宽度（Cm 对象列表）。"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.width = widths[i]
        shade_cell(cell, NAVY_HEX)
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        fmt_para(p, space_before=2, space_after=2,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
        add_rich(p, header, size=BODY_SIZE, color=WHITE, bold=True, ea=EA_HEAD)
    # 跨页时重复表头
    tr_pr = header_row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:tblHeader"))

    for r, row_data in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(row_data):
            cell = cells[i]
            cell.width = widths[i]
            if r % 2 == 1:
                shade_cell(cell, ROW_FILL)
            set_cell_margins(cell)
            p = cell.paragraphs[0]
            fmt_para(p, space_before=2, space_after=2)
            add_rich(p, value)
    set_table_borders(table, color="9DC3E6", size=6)
    return table


def setup_document():
    doc = Document()

    # 页面：A4 + 常规页边距
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.3)
    section.bottom_margin = Cm(2.3)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    # 正文默认样式：宋体 / Calibri，10.5pt，1.25 倍行距
    normal = doc.styles["Normal"]
    normal.font.name = LATIN
    normal.font.size = Pt(BODY_SIZE)
    normal.font.color.rgb = TEXT
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), LATIN)
    rfonts.set(qn("w:hAnsi"), LATIN)
    rfonts.set(qn("w:eastAsia"), EA_BODY)
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(4)
    return doc


# --------------------------------------------------------------------------- #
# 正文内容
# --------------------------------------------------------------------------- #
def build_geometry(doc):
    add_heading(doc, "1.1 几何模块（失分重灾区）", level=2)

    add_label(doc, "涉及题型")
    for item in [
        "正六边形阴影面积（割补 + 整体减空白）",
        "球内接正方体 / 大圆截面正方体",
        "圆内接三角形",
        "三角形中线长（中线长公式）",
        "解析几何不等式与圆的切线",
    ]:
        add_bullet(doc, item)

    add_label(doc, "核心错点")
    for item in [
        "**空间想象力不足**：习惯在立体图形里“硬想”，缺乏**二维截面分析**意识，"
        "不会把空间问题降维成平面三角形问题。",
        "**阴影面积缺方法**：缺少“**割补法**”与“**整体减空白**”思维，"
        "拿到图形就正面硬算，导致计算量大且易错。",
        "**公式推导不熟练**：球、正方体、外接球半径等空间几何公式只会“背”，"
        "不会从截面图中现场推导，一旦变形就束手无策。",
    ]:
        add_bullet(doc, item)


def build_combination(doc):
    add_heading(doc, "1.2 排列组合与概率 / 集合（逻辑陷阱）", level=2)

    add_label(doc, "涉及题型")
    for item in [
        "集合容斥原理",
        "机器人节点动态路径计数",
        "分组问题（均分与非均分）",
        "概率不等式的充分性判断",
    ]:
        add_bullet(doc, item)

    add_label(doc, "核心错点")
    for item in [
        "**分类讨论不严密**：分类标准不统一，容易**漏算**或**重复计数**。",
        "**分组计数不牢**：对“**分组均分需除以 A_k^k**”掌握不牢，"
        "把均分问题当成普通分组直接相乘。",
        "**路径计数硬套公式**：面对**少步数动态路径**，习惯套用组合数公式，"
        "而不愿意先**画树状图 / 穷举**验证。",
    ]:
        add_bullet(doc, item)


def build_condition(doc):
    add_heading(doc, "1.3 条件充分性判断与代数（逻辑思维）", level=2)

    add_label(doc, "涉及题型")
    for item in [
        "平均数 / 平均分变化问题",
        "均值不等式及其成立条件",
        "不定方程应用",
        "数列充分性判断（$S_n$ 与 $a_n$ 的关系）",
        "分段绝对值函数最小值",
        "植树问题中的重复计数",
    ]:
        add_bullet(doc, item)

    add_label(doc, "核心错点")
    for item in [
        "**混淆充分性边界**：分不清“**单条件充分**”与“**联合充分（C 选项）**”，"
        "常常凭感觉选 C。",
        "**忽视等号条件**：使用均值不等式时，忘记检验**等号成立条件**。",
        "**忽略边界情形**：由 $S_n$ 推导 $a_n$ 时忽略 **$n = 1$ 的边界条件**。",
        "**审题漏掉重复**：忽视“**角上重复计数**”，植树类问题直接按端点数量计算。",
    ]:
        add_bullet(doc, item)


def build_strategy(doc):
    add_heading(doc, "二、突破 60–66 分高效学习策略与周计划", level=1)

    add_heading(doc, "2.1 复习节奏：“3 + 3 + 1”模式", level=2)
    add_body(doc, "以周为单位固定节奏：**限时模考 → 深度复盘 → 专项横向复刷**，"
                  "让每一次练习都对应明确的复盘动作。")
    add_table(
        doc,
        ["时间", "安排", "关键动作"],
        [
            ["周二 / 周四", "**限时模考**",
             "50–55 分钟闭环作答一套模块练习或半套真题，不查书、不停表"],
            ["周三 / 周五", "**深度复盘与“一题三改”**",
             "逐题归因，把错题按“几何 / 排列组合 / 条件充分性”归档"],
            ["周日", "**专项横向复刷**",
             "同一题型集中重做 5–10 题，形成“题型—方法—易错点”对应表"],
        ],
        [Cm(2.9), Cm(4.7), Cm(9.0)],
    )
    add_spacer(doc, pt=6)

    add_callout(
        doc,
        "核心动作：一题三改",
        [
            "**一改思路**——这道题正确的入口在哪里？我当时为什么没看到？",
            "**二改步骤**——哪一步算错或绕远？有没有更短的推导路径？",
            "**三改陷阱**——命题人埋了什么坑（边界、等号、重复计数、截面）？",
            "要求：每道错题至少写出 **1 条可复用的“避坑口诀”**，"
            "并在周日专项复刷时验证。",
        ],
    )

    add_heading(doc, "2.2 四大破局技巧", level=2)
    for idx, item in enumerate([
        "**几何**：养成**画二维截面图**的习惯，活用“**割补法**”。"
        "球的问题先画大圆截面，正方体的问题先画对角面，"
        "把三维问题降维成三角形问题；阴影面积优先尝试“整体减空白”。",
        "**条件充分性判断**：单条件一律用**特值法找反例**；"
        "联合判断时**严禁带着条件(1)的结论去推(2)**，"
        "必须先分别确认两个条件各自不充分，才能选 C。",
        "**排列组合**：少步数的动态路径直接**画树状图 / 穷举**，不硬套公式；"
        "分组问题严格遵守“**均分除以 A_k^k**”。",
        "**考场策略**：建立 **50 秒舍弃原则**，超时难题果断标记放弃，"
        "先把基础分稳稳拿住。",
    ], start=1):
        add_numbered(doc, idx, item)
    add_spacer(doc, pt=4)

    add_callout(
        doc,
        "考场策略：50 秒舍弃原则",
        [
            "一道题思考 **50 秒**仍无清晰思路，立即标记跳过，"
            "先完成会做的题，最后统一回看。",
            "目标：**错题数控制在 3–5 题以内，总分稳定在 60–66 分**。",
        ],
        fill=BLUE_FILL,
        line_color=NAVY_HEX,
        title_color=NAVY,
    )


def build_plan(doc):
    add_heading(doc, "三、备考阶段复习规划", level=1)
    add_body(doc, "三个阶段层层递进，每个阶段都有可检验的量化目标。")
    add_table(
        doc,
        ["阶段", "核心任务", "目标要求"],
        [
            ["**阶段一**\n强化巩固",
             "专项突破**几何**与**排列组合**；重做 **2017–2021 年**真题错题，"
             "逐题写清截面图与分类依据。",
             "做到**不看答案独立推导**；几何模块正确率 ≥ 80%，"
             "分组与路径计数不再漏算、重算。"],
            ["**阶段二**\n真题刷题",
             "刷 **2022–2026 年**真题，**严格计时 55 分钟**，"
             "完全按考场节奏作答并统计错题分布。",
             "测试分数**稳定在 60+**，错题数 ≤ 5 题，"
             "条件充分性判断正确率明显提升。"],
            ["**阶段三**\n冲刺查漏",
             "复习**错题本**，重点规避**条件充分性判断**的坑点，"
             "每天保持 20–30 题的手感训练。",
             "保持**解题手感与审题速度**，错题数稳定在 **3–5 题**，"
             "总分冲击 **60–66 分**。"],
        ],
        [Cm(2.9), Cm(7.2), Cm(6.5)],
    )
    add_spacer(doc, pt=8)

    closing = doc.add_paragraph()
    fmt_para(closing, space_before=4, space_after=0,
             align=WD_ALIGN_PARAGRAPH.CENTER)
    add_rich(closing, "**目标分数：60–66 分　|　错题数：3–5 题以内　|　"
                      "核心原则：基础题不丢分，难题敢舍弃**",
             size=11, color=NAVY, ea=EA_HEAD)


def build_document():
    doc = setup_document()

    add_title(
        doc,
        "2027 考研管综数学（199）真题错题分析与高效备考方案",
        "目标分数：60–66 分　|　错题数控制在 3–5 题以内",
    )

    add_heading(doc, "一、错题深度诊断与归因分析", level=1)
    add_body(doc, "结合历年真题与近期模考数据，错题集中暴露在**几何**、"
                  "**排列组合与概率 / 集合**、**条件充分性判断与代数**三大模块。"
                  "下面按“涉及题型 → 核心错点”逐层拆解，定位失分根源。")

    build_geometry(doc)
    build_combination(doc)
    build_condition(doc)

    build_strategy(doc)
    build_plan(doc)
    return doc


def main():
    output_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "2027考研管综数学199错题分析与高效备考方案.docx",
    )
    document = build_document()
    document.save(output_path)
    print("[OK] Document generated:", output_path)


if __name__ == "__main__":
    main()
