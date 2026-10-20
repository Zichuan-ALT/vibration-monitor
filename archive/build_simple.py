# -*- coding: utf-8 -*-
"""生成《项目 8.1 振动状态监测器 — 实施方案》）。"""
import datetime
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
OUT = os.path.join(BASE, "8.1_振动状态监测器_实施方案.docx")

DARK = RGBColor(0x1F, 0x3B, 0x63)
BLACK = RGBColor(0x20, 0x20, 0x20)
ACCENT = RGBColor(0xA0, 0x3A, 0x1E)

doc = Document()
sec = doc.sections[0]
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.6)


def setfont(run, size=10.5, bold=False, ascii_font="Times New Roman",
            ea="宋体", color=BLACK):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), ascii_font)
    rf.set(qn("w:hAnsi"), ascii_font)
    rf.set(qn("w:eastAsia"), ea)


def fix_style(name, size, ea="宋体", ascii_font="Times New Roman",
              bold=False, color=BLACK):
    st = doc.styles[name]
    st.font.size = Pt(size)
    st.font.bold = bold
    st.font.name = ascii_font
    st.font.color.rgb = color
    rPr = st.element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), ascii_font)
    rf.set(qn("w:hAnsi"), ascii_font)
    rf.set(qn("w:eastAsia"), ea)
    return st


fix_style("Normal", 10.5)
doc.styles["Normal"].paragraph_format.line_spacing = 1.3
doc.styles["Normal"].paragraph_format.space_after = Pt(4)
for lvl, (sz, ea) in {1: (15, "黑体"), 2: (12, "黑体"), 3: (10.8, "黑体")}.items():
    st = fix_style(f"Heading {lvl}", sz, ea=ea, ascii_font="Arial", bold=True,
                   color=DARK)
    st.paragraph_format.space_before = Pt(11 if lvl == 1 else 8)
    st.paragraph_format.space_after = Pt(5)
    st.paragraph_format.line_spacing = 1.15
    st.paragraph_format.keep_with_next = True


def para(text="", *, size=10.5, bold=False, align="justify", ea="宋体",
         color=BLACK, indent=True, space_after=4):
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    setfont(p.add_run(text), size=size, bold=bold, ea=ea, color=color)
    return p


def rich(parts, *, size=10.5, indent=True, space_after=4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    for t, b in parts:
        setfont(p.add_run(t), size=size, bold=b)
    return p


def h(text, level=1):
    p = doc.add_heading("", level=level)
    setfont(p.add_run(text), size={1: 15, 2: 12, 3: 10.8}[level], bold=True,
            ea="黑体", ascii_font="Arial", color=DARK)
    return p


def bullets(items, *, size=10.5):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.26
        p.paragraph_format.left_indent = Pt(20)
        if isinstance(it, str):
            setfont(p.add_run(it), size=size)
        elif isinstance(it, tuple):
            setfont(p.add_run(it[0]), size=size, bold=it[1])
        else:
            for t, b in it:
                setfont(p.add_run(t), size=size, bold=b)


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def set_cell(cell, text, *, size=8.6, bold=False, align="left", ea="宋体"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER}[align]
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.1
    setfont(p.add_run(text), size=size, bold=bold, ea=ea)


def table(caption, headers, rows, widths, *, size=8.6, align_map=None,
          header_fill="DEEBF7"):
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(6)
    cp.paragraph_format.space_after = Pt(2)
    cp.paragraph_format.keep_with_next = True
    if caption:
        setfont(cp.add_run(caption), size=9.5, bold=True, ea="黑体", color=DARK)
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, htxt in enumerate(headers):
        c = t.rows[0].cells[i]
        set_cell(c, htxt, size=size, bold=True, align="center", ea="黑体")
        shade(c, header_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            a = "center" if (align_map and align_map[i] == "c") else "left"
            set_cell(cells[i], str(v), size=size, align=a)
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = widths[i]
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def plain_table(rows, widths, *, size=8.8, label_fill="DEEBF7"):
    """两列“标签 / 内容”表，用于步骤卡。"""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for k, v in rows:
        cells = t.add_row().cells
        set_cell(cells[0], k, size=size, bold=True, align="center", ea="黑体")
        shade(cells[0], label_fill)
        set_cell(cells[1], v, size=size)
    for r in t.rows:
        r.cells[0].width = widths[0]
        r.cells[1].width = widths[1]
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def figure(fname, caption, width_cm):
    doc.add_picture(os.path.join(FIG, fname), width=Cm(width_cm))
    p = doc.paragraphs[-1]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(8)
    setfont(cp.add_run(caption), size=9, ea="黑体",
            color=RGBColor(0x50, 0x50, 0x50))


def code_block(lines):
    for i, ln in enumerate(lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(4) if i == 0 else Pt(0)
        p.paragraph_format.left_indent = Pt(12)
        p.paragraph_format.line_spacing = 1.02
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "F4F6F8")
        pPr.append(shd)
        setfont(p.add_run(ln), size=8.4, ascii_font="Consolas", ea="宋体")
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# 页脚
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
setfont(fp.add_run("项目 8.1 振动状态监测器 · 实施方案　第 "), size=8)
r = fp.add_run()
f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
r._r.append(f1); r._r.append(it); r._r.append(f2)
setfont(r, size=8)
setfont(fp.add_run(" 页"), size=8)

TODAY = datetime.date.today()

# ================================================================= 封面
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("利莫瑞克大学 × LICHPU　工业评估（15 ECTS）"), size=11.5,
        bold=True, ea="黑体", color=ACCENT)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
setfont(p.add_run("智慧校园生活实验室 · 主题 8 智能设施与预测性维护"), size=11,
        ea="黑体", color=RGBColor(0x55, 0x5B, 0x64))
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("项目 8.1　振动状态监测器"), size=24, bold=True, ea="黑体",
        color=DARK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
setfont(p.add_run("振动状态监测器 实施方案"), size=15, bold=True, ea="黑体",
        color=ACCENT)

table("", ["项目", "内容"],
      [["文件性质", "本项目单节点样机的完整实施方案，覆盖选型、耗材、接线、施工、"
                "固件、服务端、校准与验证"],
       ["技术路线", "ESP32 单板 + 振动 / 温度两路数字传感 + Node-RED 仪表板"],
       ["硬件成本", "¥211（含耗材）；加备件约 ¥232；焊接耗材由实验室提供时约 ¥179"],
       ["实施工时", "约 13 ~ 17 个工作日（不含报告与答辩）"],
       ["核心内容", "13 个施工步骤，每步给出验收标准与常见问题；"
                "12 项需求与 12 条验收标准；逐项耗材清单"],
       ["版本 / 日期", f"V2.1　{TODAY.strftime('%Y-%m-%d')}"]],
      [Cm(2.8), Cm(13.2)], size=9.5, align_map=["c", "l"])

# ================================================================= 一
h("一、设计思路与总览", 1)
h("1.1 设计思路", 2)
para("本项目的评估权重中，技术设计与实现（C）占 25%，而测试、验证与数据分析（D）占 20%、"
     "报告（E）占 15%、专业执行（B）占 15%。也就是说，四分之三的分数来自"
     "“系统做完、测过、记录清楚”。本方案因此把工程复杂度压到最低：只用数字接口传感器、"
     "只用一块主控、服务端复用现有笔记本；把省下的时间全部投入到校准、对照实验与数据分析上，"
     "同时不触碰手册规定的任何硬性下限。", indent=True)

h("1.2 关键设计选择", 2)
table("表 1　关键设计选择与理由", ["设计环节", "本方案的选择", "理由"],
      [["主控", "ESP32-WROOM-32 单板（内置 Wi-Fi）",
        "一块板同时完成采集、计算与联网，省掉模块间通信与额外调试"],
       ["传感输入", "振动 + 温度两路，均为数字接口",
        "满足“至少两个有意义的数据输入”要求；数字接口免去模拟前端与 ADC 调试"],
       ["本地界面", "RGB LED + 蜂鸣器 + 1 个按键", "本地指示够用，界面工作量交给仪表板"],
       ["边缘算法", "FFT + 速度有效值 + 基线自学习 + 三级判别",
        "在本地给出结论而非转发原始数值，满足对手边处理深度的要求"],
       ["服务端", "Node-RED 单进程 + SQLite（跑在笔记本上）",
        "一个进程同时充当 MQTT Broker、历史库与仪表板，零硬件采购"],
       ["供电", "ESP32 用 USB、风机用 12 V 适配器，两路独立",
        "避免电机噪声串入数字电源——这是数据异常最常见的原因"],
       ["硬件成本", "¥211（含耗材），加备件约 ¥232", "见第二章逐项清单"],
       ["实施工时", "约 13 ~ 17 个工作日", "把时间集中投入到校准、对照实验与数据分析"]],
      [Cm(2.2), Cm(6.4), Cm(7.4)], size=8.6, align_map=["l", "l", "l"])

h("1.3 系统总览", 2)
figure("figS1_arch.png", "图 S1　系统架构", 15.8)
rich([("关键点：", True),
      ("ADXL345 与 DS18B20 都是数字接口（SPI 与 1-Wire），因此整个系统不需要 ADC、"
       "不需要模拟前端、也不需要信号调理电路。这既省掉了元件与调试，也正好绕开"
       "ESP32 片内 ADC 噪声偏大的短板。2048 点 FFT 用 ESP-DSP 库约需 1 ~ 2 ms，"
       "相对 0.64 s 的采样窗口完全不是瓶颈。", False)])

h("1.4 总览", 2)
table("表 2　方案总览", ["项目", "内容"],
      [["硬件成本", "¥211（主要器件 ¥72 + 电子耗材 ¥57 + 结构与台架耗材 ¥82）；"
                "按 10% 计备件与损耗约 ¥232；若焊接耗材与标准件由实验室提供，约 ¥179"],
       ["服务端成本", "¥0（使用现有笔记本 + Node-RED + SQLite，不采购服务器）"],
       ["实施工时", "约 13 ~ 17 个工作日（不含报告、海报与答辩准备）"],
       ["里程碑", "第 5 ~ 6 周即可得到可演示的完整系统（采集 → 判别 → 上报 → 仪表板）"],
       ["交付物", "节点样机 1 台、固件源码 1 套、Node-RED 流程 1 套、"
               "实验数据集 1 份、校准与测试记录 1 套"]],
      [Cm(2.6), Cm(13.4)], size=8.8, align_map=["c", "l"])

# ================================================================= 二
h("二、物料与耗材清单", 1)
para("本章按“主要器件 / 电子耗材 / 结构与台架耗材”三类逐项列出，"
     "价格为中国零售参考价（含税、不含运费）。", indent=True)

h("2.1 主要器件", 2)
table("表 3　主要器件清单", ["#", "物料", "规格 / 说明", "数量", "单价 ¥", "小计 ¥"],
      [["1", "主控板", "ESP32-WROOM-32 开发板（含 USB 数据线）", "1", "25", "25"],
       ["2", "振动传感器", "ADXL345 模块（SPI / I²C，3.3 V）", "1", "12", "12"],
       ["3", "温度传感器", "DS18B20 防水探头（1-Wire，带 1 m 线）", "1", "8", "8"],
       ["4", "存储卡座", "microSD 卡座模块（SPI，3.3 V）", "1", "5", "5"],
       ["5", "存储卡", "microSD 8 GB Class10", "1", "15", "15"],
       ["6", "状态指示", "RGB LED 模块，或三色 LED + 3 个电阻", "1", "3", "3"],
       ["7", "蜂鸣器", "有源蜂鸣器 5 V（带驱动板）", "1", "2", "2"],
       ["8", "按键", "12×12 轻触按键（带帽）", "2", "1", "2"],
       ["", "小计", "", "", "", "72"]],
      [Cm(0.9), Cm(2.4), Cm(7.6), Cm(1.3), Cm(1.7), Cm(1.7)], size=8.6,
      align_map=["c", "l", "l", "c", "c", "c"])

h("2.2 电子耗材", 2)
table("表 4　电子耗材清单", ["#", "物料", "规格 / 说明", "数量", "单价 ¥", "小计 ¥"],
      [["9", "杜邦线", "母-母 40 根 + 公-母 20 根", "1 套", "8", "8"],
       ["10", "单芯导线", "0.5 mm²，红黑各 3 m", "1 套", "6", "6"],
       ["11", "洞洞板", "5×7 cm 双面镀锡", "2", "2", "4"],
       ["12", "排针 / 排母", "2.54 mm 混装", "1 套", "5", "5"],
       ["13", "焊锡丝", "0.6 mm，含锡 63 %", "1 小卷", "15", "15"],
       ["14", "助焊剂 / 吸锡带", "免洗助焊剂或松香", "1 套", "8", "8"],
       ["15", "热缩管", "直径 2 / 3 / 5 mm 混装", "1 套", "6", "6"],
       ["16", "扎带与魔术贴", "小号尼龙扎带", "1 套", "5", "5"],
       ["", "小计", "第 13 ~ 16 项实验室常备，若可借用可省 ¥34", "", "", "57"]],
      [Cm(0.9), Cm(2.4), Cm(7.6), Cm(1.3), Cm(1.7), Cm(1.7)], size=8.6,
      align_map=["c", "l", "l", "c", "c", "c"])

h("2.3 结构与台架耗材", 2)
table("表 5　结构与台架耗材清单", ["#", "物料", "规格 / 说明", "数量", "单价 ¥", "小计 ¥"],
      [["17", "塑料接线盒", "100×68×50 mm", "2", "5", "10"],
       ["18", "底板", "木板或铝条 200×100×10 mm", "1", "10", "10"],
       ["19", "M3 紧固件套装", "螺丝 / 螺母 / 垫片混装盒", "1", "10", "10"],
       ["20", "防松垫圈", "M3 弹簧垫圈", "1 包", "3", "3"],
       ["21", "反光标记", "白色标签纸或反光贴（用于转速观察）", "1 张", "3", "3"],
       ["22", "配重件", "5 g / 10 g 配重块各若干（也可用硬币、螺母）", "1 套", "5", "5"],
       ["23", "热熔胶棒与双面胶", "固定走线用", "1 套", "6", "6"],
       ["24", "演示风机", "12 V 直流风机（电脑机箱风扇 80/120 mm）", "1", "15", "15"],
       ["25", "电源适配器", "12 V / 1 A，供风机", "1", "20", "20"],
       ["", "小计", "第 19、20、23 项实验室常备，若可借用可省 ¥19", "", "", "82"]],
      [Cm(0.9), Cm(2.4), Cm(7.6), Cm(1.3), Cm(1.7), Cm(1.7)], size=8.6,
      align_map=["c", "l", "l", "c", "c", "c"])

h("2.4 需要准备的工具", 2)
bullets([
    "电烙铁（含烙铁架、海绵）与焊锡 —— 焊接排针与蜂鸣器驱动；",
    "万用表 —— 每一步接线后先量通断与电压，再上电；",
    "螺丝刀套装、剥线钳、斜口钳；",
    "热熔胶枪 —— 固定走线与传感器；",
    "手持测振仪与手持转速表 —— 校准比对用，**必须先确认实验室能否借到**"
    "（这是风险 R-02 的前提）。",
], size=10)

h("2.5 替代型号与采购要点", 2)
table("表 6　替代型号与采购要点", ["物料", "首选", "可用替代", "采购要点"],
      [["主控板", "ESP32-WROOM-32 DevKit", "ESP32-S3、ESP32-C3（引脚需重排）",
        "**买两块**，一块备用；选带 USB-C 与稳压芯片的版本"],
       ["振动传感器", "ADXL345 模块", "ICM-42688-P、LIS3DH（需改驱动）",
        "选**模块**不选裸芯片；确认支持 SPI（有的只引 I²C）"],
       ["温度传感器", "DS18B20 防水探头", "SHT31（I²C）、NTC + 分压",
        "买带线的探头，方便贴到风机外壳"],
       ["存储", "microSD SPI 模块", "直接用 ESP32 片内 SPIFFS（容量小）",
        "注意模块是 3.3 V 供电；卡用 Class10 以上"],
       ["风机", "12 V 机箱风扇", "24 V 直流电机 + 扇叶",
        "机箱风扇最省事：转速稳定在 1 200 ~ 1 500 r/min，正好落在分析带内"]],
      [Cm(2.0), Cm(3.6), Cm(4.2), Cm(6.2)], size=8.4,
      align_map=["l", "l", "l", "l"])

h("2.6 成本汇总", 2)
table("表 7　成本汇总", ["项目", "金额 ¥", "说明"],
      [["主要器件（表 3）", "72", "ESP32、两路传感器、存储、指示与按键"],
       ["电子耗材（表 4）", "57", "导线、洞洞板、焊接与绝缘耗材"],
       ["结构与台架耗材（表 5）", "82", "外壳、底板、紧固件、配重、风机与适配器"],
       ["小计", "211", "4 人分摊约 ¥53 / 人"],
       ["备件与损耗（按 10 % 计）", "21", "建议备一块 ESP32 与少量导线"],
       ["合计（含备件）", "232", "—"],
       ["若焊接耗材与标准件由实验室提供", "179", "可省第 13 ~ 16、19、20、23 项"]],
      [Cm(5.0), Cm(2.2), Cm(8.8)], size=8.8, align_map=["l", "c", "l"])
rich([("与上一版的差异说明：", True),
      ("上一版（V1.0）报出的 ¥155 中，“外壳与线材 / 台架材料 / 杂项”三行是打包估算（合计 ¥45）。"
       "本手册把这部分展开为 15 个具体条目，因此耗材由 ¥45 增至 ¥139（表 4 + 表 5），"
       "器件部分基本不变（¥70 → ¥72）。差额 ¥56 全部落在耗材上，"
       "主要是焊接材料、紧固与防松件、配重件和 12 V 适配器。"
       "建议按 ¥211 编列预算，并如实说明其中约 ¥53 可由实验室常备品替代。", False)])

# ================================================================= 三
h("三、接线与装配", 1)
h("3.1 接线图", 2)
figure("figS_wiring.png", "图 S2　接线示意图（④ 为本地指示与按键）", 15.8)

h("3.2 引脚分配", 2)
table("表 8　ESP32 引脚分配", ["引脚", "功能", "连接对象", "说明"],
      [["GPIO18", "SPI SCK", "ADXL345 + microSD", "共用时钟"],
       ["GPIO19", "SPI MISO", "ADXL345 + microSD", "共用数据输入"],
       ["GPIO23", "SPI MOSI", "ADXL345 + microSD", "共用数据输出"],
       ["GPIO5", "SPI CS", "ADXL345", "片选，软件控制"],
       ["GPIO2", "SPI CS", "microSD 模块", "片选，与 ADXL345 分时复用总线"],
       ["GPIO4", "1-Wire", "DS18B20", "4.7 kΩ 上拉到 3V3"],
       ["GPIO25 / 26 / 27", "GPIO 输出", "RGB LED", "各串 1 kΩ 限流"],
       ["GPIO14", "PWM 输出", "蜂鸣器", "经 NPN 三极管驱动"],
       ["GPIO0", "输入（内部上拉）", "按键", "短按静音，长按重新学习基线"],
       ["GPIO1 / 3", "USB 串口", "调试与烧录", "板载 USB 完成"],
       ["3V3 / GND", "电源", "全部模块", "**一律 3.3 V，禁止接 5 V**"]],
      [Cm(3.0), Cm(3.0), Cm(3.6), Cm(6.4)], size=8.6,
      align_map=["c", "l", "l", "l"])

h("3.3 供电与接地", 2)
bullets([
    "ESP32 由 USB 充电器（5 V）供电；12 V 直流风机由独立适配器供电。**两路分开**，"
    "避免电机噪声串入数字电源——这是本项目最容易忽视又最容易导致数据异常的问题。",
    "所有模块的 GND 与 ESP32 共地。",
    "传感器与主控之间的连线尽量短（≤ 20 cm），并远离风机电源线。",
])

h("3.4 装配顺序", 2)
table("表 9　装配顺序与注意事项", ["顺序", "动作", "注意事项"],
      [["1", "把排针焊到 ADXL345 / microSD 模块上", "先焊排针再上电，避免虚焊反复排查"],
       ["2", "洞洞板上焊好 ESP32 与各模块的电源总线（3V3 / GND）",
        "电源线用红色与黑色区分，先只接电源不接信号"],
       ["3", "焊接 RGB LED、蜂鸣器驱动三极管与按键", "LED 各串 1 kΩ；按键一端接 GND"],
       ["4", "用杜邦线接 ADXL345（先只接这一路）", "接完先量通断，再上电验证"],
       ["5", "依次接入 DS18B20、microSD", "每接一路都要重新做一次自检"],
       ["6", "装入塑料接线盒，热熔胶固定走线", "传感器要伸出盒外，便于贴到风机上"],
       ["7", "把风机拧到底板上，传感器用螺栓或粘接固定在风机外壳",
        "**刚性安装**，不要用双面胶加海绵——会衰减高频"]],
      [Cm(1.3), Cm(7.5), Cm(7.2)], size=8.6, align_map=["c", "l", "l"])

# ================================================================= 四
h("四、分步实施（13 步）", 1)
para("每一步都遵循同一个原则：**只增加一个变量，做完立即验证并记录**。"
     "表格中的“耗时”为单人估算，两人协作可压缩约三分之一。", indent=True)

steps = [
    ("表 10　步骤 0：采购与工具准备（0.5 天）",
     "备齐第二章全部物料，安装开发与服务端软件",
     "① 按表 3 ~ 表 5 下单（ESP32 买两块）；② 安装 Arduino IDE 2.x 并在板管理器中"
     "添加 ESP32 支持包；③ 安装 Node.js LTS，执行 npm install -g --unsafe-perm node-red；"
     "④ 在 Node-RED 的“节点管理”中安装 node-red-dashboard、node-red-node-sqlite、"
     "node-red-contrib-aedes 三个模块",
     "Arduino IDE 能编译 ESP32 的 Blink 示例；浏览器打开 http://localhost:1880 看到编辑器；"
     "三个 palette 出现在左侧节点栏",
     "别一次买齐所有替代型号，先按表 3 下单；aedes 模块负责充当 MQTT Broker，"
     "装了它就不必再单独安装 Mosquitto"),
    ("表 11　步骤 1：点亮最小系统（0.5 天）",
     "确认工具链、驱动、串口全部正常",
     "烧录 Blink + 串口打印程序，串口监视器波特率 115200",
     "板载 LED 闪烁；串口能看到周期输出",
     "这一步同时排除三个坑：USB 驱动未装、开发板型号选错、串口被占用。"
     "若上传失败，按住 BOOT 键再点上传"),
    ("表 12　步骤 2：接通振动传感器（1 天）",
     "把 ADXL345 读通，拿到可信的加速度原始值",
     "按表 8 接线 → 先读器件 ID（寄存器 0x00，应返回 0xE5）→ 再读三轴原始值",
     "ID 正确；静止时某一轴约为 ±1 g（取决于安装方向）；轻敲桌面数值明显跳变",
     "模块丝印在 SPI 与 I²C 下含义不同（SDO/SDA 在 SPI 下是 MISO/MOSI）。"
     "**卡住就先用 I²C 跑通**（CS 接高），确认芯片活着再切 SPI，别在接线上耗两天"),
    ("表 13　步骤 3：接通温度传感器（0.5 天）",
     "读到可信的表面温度",
     "GPIO4 接数据线，4.7 kΩ 上拉到 3V3；用 DallasTemperature 库读取",
     "室温读数合理；手指捏住探头，温度在数秒内上升 1 ~ 2 ℃",
     "若读到 −127 ℃，说明时序或上拉有问题；先只接一个 DS18B20，"
     "多探头需要先搜索 ROM 地址"),
    ("表 14　步骤 4：接通 microSD 存储（0.5 天）",
     "能写能读，为断网缓存做准备",
     "与 ADXL345 共用 SPI 总线，CS 用 GPIO2；用 SD 库挂载并写一个 CSV 文件再读回",
     "文件创建成功，内容读回一致",
     "microSD 模块必须 3.3 V 供电；文件名用 8.3 短名兼容性最好；"
     "挂载失败先降低 SPI 时钟到 4 MHz 试试"),
    ("表 15　步骤 5：完成定时采样与 FFT（2 ~ 3 天）",
     "把时域波形变成可信的频谱——这是全项目的技术核心",
     "① 用微秒级定时循环实现 3.2 kSPS × 2048 点采样；② 去均值 + Hanning 窗；"
     "③ 用 arduinoFFT 做 2048 点实数 FFT；④ 换算幅值谱",
     "采 1 s 统计点数为 3200 ± 16；把传感器贴在小音箱振膜上放 100 Hz 正弦，"
     "谱峰出现在 100 Hz ± 1 %",
     "**不要用 delay() 循环采样**，采样率会飘、频谱全假；"
     "把采样放在高优先级任务并绑定到 ESP32 的 core 1，网络留在 core 0，"
     "否则 Wi-Fi 会打断采样节拍"),
    ("表 16　步骤 6：实现特征计算（1 天）",
     "把频谱与时域数据压缩成几个可判读的指标",
     "计算 RMS、峰值、速度有效值（10 ~ 1000 Hz 积分）以及基频带（15 ~ 35 Hz）"
     "内的最大幅值及其频率",
     "各指标随时间平稳；用手挡风道时 RMS 与温度同时上升",
     "没有转速传感器时，**不要去找精确的 1×**，改用“基频带内最大峰值”作为指标，"
     "这样转速小范围波动时结果依然稳定"),
    ("表 17　步骤 7：基线自学习与三级判别（1 ~ 2 天）",
     "让节点自己判断设备状态",
     "① 开机连续采集 60 s，算振动 RMS 的 μ 与 σ；② 按 μ+3σ / μ+6σ 与温度增量"
     "+8 ℃ / +15 ℃ 分级；③ 加迟滞（连续 3 帧确认报警、连续 5 帧确认恢复）；"
     "④ 驱动 RGB LED 与蜂鸣器",
     "叶轮贴配重后状态由正常升到注意或报警，撕掉后能自动恢复；"
     "正常工况连续运行 2 h 内假阳性不超过 1 次",
     "基线只在“连续 30 min 处于正常状态”时缓慢更新，"
     "否则会把已经劣化的状态学成新的正常，导致漏报"),
    ("表 18　步骤 8：接入 MQTT 上报（1 天）",
     "把结果送到笔记本",
     "用 PubSubClient 组 JSON 并发到笔记本 IP 的 1883 端口，周期 10 s，事件即时上报",
     "在笔记本上用 MQTTX 订阅 smartcampus/vib/# 能看到完整 JSON；"
     "断网后节点继续运行并在恢复后补传",
     "**第一件事就加 client.setBufferSize(1024);** PubSubClient 默认缓冲区只有 256 字节，"
     "JSON 超长会静默发送失败，查半天查不出来。笔记本要在路由器上做 DHCP 保留固定 IP"),
    ("表 19　步骤 9：搭建 Node-RED 服务端与仪表板（1 ~ 2 天）",
     "得到实时 + 历史 + 报警的可视化界面",
     "按表 24 搭建流程：aedes broker → mqtt in → json → function（状态与建议）→ "
     "dashboard 仪表与图表 + SQLite 存历史 + switch 报警",
     "浏览器打开 http://localhost:1880/ui 看到实时刷新的状态卡、趋势曲线与报警列表；"
     "重启 Node-RED 后历史数据仍在",
     "SQLite 节点要显式建表（见 6.3），否则写入会静默失败；"
     "ui_chart 默认只保留内存中的点数，长趋势要靠 SQLite 查询回填"),
    ("表 20　步骤 10：完成校准（1 天）",
     "让测量结果可追溯、可解释",
     "重力翻转法标定振动灵敏度与零偏；冰水浴与温水比对温度；"
     "与手持测振仪同点同时比对幅值",
     "见第十章表 25 的判据；全部记录填入台账「6-测试记录」",
     "没有振动台时，报告里必须写明这是“比对”而不是“绝对溯源”，并给出不确定度预算。"
     "**先确认实验室能借到手持测振仪**，否则这条判据要降级为相对趋势验证"),
    ("表 21　步骤 11：搭建演示台架并制造工况（1 天）",
     "能在台架上稳定复现三类工况变化",
     "把风机固定到底板；准备 2 g / 5 g 配重片、可松动的底座螺丝、"
     "可遮挡 30 % / 60 % 的纸板",
     "每种工况施加方式可重复，施加前后特征值有明确差异",
     "配重用胶带贴在叶轮同一位置并记录质量；松螺丝要记录圈数；"
     "遮挡板要记录面积比例——**可复现性是评估组成 D 的核心**"),
    ("表 22　步骤 12：对照实验与数据分析（3 ~ 5 天）",
     "积累足以支撑结论的实验证据",
     "按第十一章表 26 的实验矩阵逐项执行，每种工况按规定次数重复，"
     "记录全部特征值与原始 2048 点数据",
     "每个工况的样本量达标；用箱线图与 ROC 曲线评估各判据的检出能力，"
     "并报告假阳性率",
     "数据当场记录，不要事后补；原始数据统一命名并存到 Git 仓库的 evidence 目录"),
]
for cap, goal, doing, accept, pit in steps:
    table(cap, ["项目", "内容"],
          [["目标", goal], ["做什么", doing], ["验收标准", accept], ["常见问题", pit]],
          [Cm(2.2), Cm(13.8)], size=8.6, align_map=["c", "l"])

# ================================================================= 五
h("五、固件实现要点", 1)
h("5.1 程序结构", 2)
figure("figS2_flow.png", "图 S3　固件主流程", 10.4)
bullets([
    "**双核分工**：采样任务用 xTaskCreatePinnedToCore 绑定到 core 1 并设高优先级，"
    "Wi-Fi 与 MQTT 留在 core 0。否则网络中断会打乱采样节拍，频谱失真。",
    "**依赖库**：Adafruit_ADXL345、OneWire + DallasTemperature、arduinoFFT（先用它跑通，"
    "再考虑换 ESP-DSP 提速）、PubSubClient、SD、ArduinoJson。",
    "**调试开关**：留一个串口命令（如输入 r 重新学习基线、输入 d 打印一帧特征），"
    "现场调试效率会高很多。",
])

h("5.2 关键参数配置", 2)
table("表 23　固件关键参数", ["参数", "取值", "说明"],
      [["采样率 fs", "3200 Hz", "与 ADXL345 的 ODR 联动"],
       ["帧长 N", "2048 点", "Δf = fs / N = 1.5625 Hz"],
       ["帧周期", "1 s", "保证报警端到端延迟 ≤ 5 s"],
       ["采样时钟", "312 μs / 点", "用 micros() 或硬件定时器"],
       ["基频带", "15 ~ 35 Hz", "对应 900 ~ 2100 r/min，用于取基频峰值"],
       ["速度积分区间", "10 ~ 1000 Hz", "积分发散的频率在区间外剔除"],
       ["基线自学习时长", "60 s", "开机后自动执行"],
       ["基线更新条件", "连续 30 min 正常", "防止把故障学成正常"],
       ["注意 / 报警阈值", "μ + 3σ / μ + 6σ", "振动 RMS"],
       ["温度注意 / 报警", "+8 ℃ / +15 ℃", "相对基线的增量"],
       ["报警确认 / 恢复", "连续 3 帧 / 5 帧", "迟滞，抑制抖动"],
       ["MQTT 上报周期", "10 s", "事件即时上报"]],
      [Cm(3.6), Cm(4.0), Cm(8.4)], size=8.6, align_map=["l", "c", "l"])

h("5.3 关键代码要点", 2)
rich([("① 定时采样（ESP32 双核，避开 Wi-Fi 干扰）", True)], indent=False, space_after=2)
code_block([
    "const int N = 2048;  const float FS = 3200.0;  const uint32_t T_US = 312;",
    "void sampleTask(void *p) {",
    "  uint32_t next = micros();",
    "  for (int i = 0; i < N; i++) {",
    "    while ((int32_t)(micros() - next) < 0) { }   // 等待到下一个采样时刻",
    "    next += T_US;",
    "    buf[i] = adxl.getX();                        // 读单轴原始值",
    "  }",
    "  vTaskDelete(NULL);",
    "}",
    "// 在 setup() 中：采样任务绑到 core 1，优先级 3",
    "xTaskCreatePinnedToCore(sampleTask, \"sampling\", 4096, NULL, 3, NULL, 1);",
])
rich([("② FFT（arduinoFFT，先跑通再优化）", True)], indent=False, space_after=2)
code_block([
    "double vReal[N], vImag[N];",
    "double mean = 0; for (int i = 0; i < N; i++) mean += buf[i]; mean /= N;",
    "for (int i = 0; i < N; i++) { vReal[i] = buf[i] - mean; vImag[i] = 0; }",
    "FFT.Windowing(vReal, N, FFT_WIN_TYP_HANN, FFT_FORWARD);   // Hanning 窗",
    "FFT.Compute(vReal, vImag, N, FFT_FORWARD);",
    "FFT.ComplexToMagnitude(vReal, vImag, N);",
    "double fPeak = FFT.MajorPeak(vReal, N, FS);               // 主频",
])
rich([("③ 速度有效值（10 ~ 1000 Hz 积分，单位 mm/s）", True)], indent=False,
     space_after=2)
code_block([
    "double s = 0;",
    "for (int k = 1; k < N / 2; k++) {",
    "  double f = k * FS / N;",
    "  if (f < 10 || f > 1000) continue;",
    "  double a = vReal[k] / (N / 2.0) * 2.0;   // Hanning 窗需乘幅值恢复系数",
    "  double v = a / (2 * PI * f);             // 加速度谱积分成速度谱",
    "  s += v * v;",
    "}",
    "double v_rms = sqrt(s / 2) * 1000.0;       // m/s → mm/s",
])
rich([("④ MQTT 发送（注意缓冲区）", True)], indent=False, space_after=2)
code_block([
    "client.setServer(BROKER_IP, 1883);",
    "client.setBufferSize(1024);   // 默认仅 256 字节，JSON 会静默发送失败",
])

# ================================================================= 六
h("六、服务端实现（Node-RED）", 1)
h("6.1 组件与配置", 2)
table("表 24　Node-RED 组件与配置要点", ["顺序", "节点", "配置要点"],
      [["1", "aedes broker", "端口 1883，允许匿名连接（仅限校园局域网，勿暴露公网）"],
       ["2", "mqtt in", "主题 smartcampus/vib/+/state，QoS 1"],
       ["3", "json", "解析 payload 为对象"],
       ["4", "function（状态与建议）", "依据 score 与温度增量生成 state 与 advice 文本"],
       ["5", "ui_gauge × 3", "振动 RMS、速度有效值、表面温度三张状态卡"],
       ["6", "ui_chart × 2", "24 h 趋势曲线（数据来自 SQLite）；频谱图（订阅 /spectrum）"],
       ["7", "sqlite（读写）", "写入 readings 表；趋势图通过查询回填历史"],
       ["8", "switch（报警）", "state 非 normal 时触发提示，并写入报警表"],
       ["9", "ui_text（报警与建议）", "分级列表，每条报警附一条维护建议"]],
      [Cm(1.3), Cm(4.2), Cm(10.5)], size=8.6, align_map=["c", "l", "l"])

h("6.2 数据库表结构", 2)
code_block([
    "CREATE TABLE readings (",
    "  ts     TEXT,     -- ISO8601 时间戳",
    "  rms    REAL,     -- 振动加速度有效值 [g]",
    "  v_rms  REAL,     -- 速度有效值 [mm/s]",
    "  temp   REAL,     -- 表面温度 [℃]",
    "  fpeak  REAL,     -- 基频带内最大峰值频率 [Hz]",
    "  state  TEXT,     -- normal | attention | alarm",
    "  score  REAL      -- 偏离度",
    ");",
    "CREATE INDEX idx_ts ON readings(ts);",
])

h("6.3 仪表板", 2)
figure("fig4_dashboard.png", "图 S4　仪表板线框（Node-RED dashboard 实现）", 15.8)
bullets([
    "实时区：状态指示灯 + 振动 RMS、速度有效值、表面温度三张状态卡；",
    "趋势区：近 24 h 曲线并叠加注意 / 报警阈值线；",
    "频谱区：2048 点幅值谱，标注基频带；",
    "报警区：分级列表，每条附一条维护建议；",
    "历史：SQLite 持久化满足“历史可查 ≥ 4 周”；节点侧 microSD 同时留底。",
])

# ================================================================= 七
h("七、校准与测试记录", 1)
table("表 25　校准步骤与判据", ["项目", "方法", "判据"],
      [["振动灵敏度与零偏", "重力翻转法：传感器六个朝向各朝下，各采样 10 s，最小二乘拟合",
        "拟合残差 ≤ 5 %"],
       ["振动幅值（比对）", "与手持测振仪同点同时测量 3 个转速点", "偏差 ≤ 10 %"],
       ["温度", "冰水浴（0 ℃）与温水（约 45 ℃）与标准温度计比对", "偏差 ≤ 0.5 ℃"],
       ["采样率", "连续采集 1 s，统计实际点数", "3200 ± 16"],
       ["频率准确度", "小音箱振膜 + 信号发生器输出 100 Hz 正弦", "谱峰误差 ≤ 1 %"],
       ["重复性", "同一工况连续测量 10 次", "变异系数 ≤ 3 %"]],
      [Cm(3.0), Cm(8.2), Cm(4.8)], size=8.6, align_map=["l", "l", "l"])
bullets([
    "每次测量当场填写台账「6-测试记录」，并保存原始数据文件；",
    "校准系数写入固件配置或 SD 卡配置文件，更换传感器后可重新标定；",
    "所有结论附不确定度说明，避免对测量能力作过度断言。",
])

# ================================================================= 八
h("八、演示台架与工况实验", 1)
h("8.1 台架搭法", 2)
bullets([
    "把 12 V 直流风机用 4 颗螺丝拧在木板或铝条上，底板再固定到桌面；",
    "传感器用螺栓或粘接固定在风机外壳靠近轴承的位置（**刚性安装**）；",
    "不平衡：在叶轮同一位置贴配重片（2 g 与 5 g 两档），记录质量与位置；",
    "结构松动：松开固定风机的一颗螺丝 1/4 圈或 1/2 圈，记录圈数；",
    "负载变化：用硬纸板遮挡进风口 30 % 与 60 %，记录遮挡比例。",
])
h("8.2 工况实验矩阵", 2)
table("表 26　工况变化对照实验矩阵", ["工况", "施加方式", "预期特征变化", "重复次数"],
      [["正常基线", "原始状态，稳定运行", "各指标稳定在 ±2σ 内", "5"],
       ["转子不平衡", "叶轮贴 2 g / 5 g 配重片", "基频带幅值与 RMS 上升", "5 / 档"],
       ["结构松动", "松开底座螺丝 1/4 圈、1/2 圈", "谐波成分上升，转速波动增大", "3 / 档"],
       ["负载变化", "遮挡进风口 30 % / 60 %", "RMS 与温度同时上升", "3 / 档"],
       ["转速变化", "调压至 9 V / 12 V", "基频带峰值随转速移动（验证归一化必要性）", "3"]],
      [Cm(2.2), Cm(4.8), Cm(5.4), Cm(3.6)], size=8.6,
      align_map=["l", "l", "l", "c"])
para("每条记录须包含：时间戳、转速、环境温度、施加参数、全部特征值、"
     "原始 2048 点数据文件名。按工况分组统计均值与标准差，"
     "用箱线图与 ROC 曲线评估各判据的检出能力，并报告假阳性率。", indent=True)

# ================================================================= 九
h("九、需求与验收标准", 1)
para("本章给出本方案自身的功能性需求、非功能性需求与可测量的验收标准，"
     "作为施工、调试与最终验证的依据。第 2 列为优先级：M 必须 / S 应该 / C 可以。", indent=True)
table("表 27　功能性需求", ["编号", "需求描述", "优先级", "验证方法"],
      [["FR-01", "采集振动（SPI，≥ 3.2 kSPS，帧长 2048 点）与温度两路数据", "M",
        "数据记录与仪表板检查"],
       ["FR-02", "计算时域特征：RMS、峰值、速度有效值（10 ~ 1000 Hz 积分）", "M",
        "注入已知信号比对"],
       ["FR-03", "完成 2048 点 FFT，给出主频与基频带（15 ~ 35 Hz）内峰值", "M",
        "音箱振膜 + 100 Hz 正弦验证"],
       ["FR-04", "支持开机基线自学习，并给出相对基线的偏离度", "M", "与离线计算比对"],
       ["FR-05", "输出“正常 / 注意 / 报警”三级状态与维护建议", "M",
        "人工制造工况变化验证"],
       ["FR-06", "通过 MQTT 上报特征值，周期 ≤ 10 s，事件即时上报", "M",
        "抓包与端到端计时"],
       ["FR-07", "断网时本地缓存，网络恢复后补传，不丢数据", "S", "拔网线测试"],
       ["FR-08", "提供中英双语仪表板，含实时值、历史趋势与频谱", "M",
        "界面走查 + 用户测试"],
       ["FR-09", "提供本地指示（RGB LED + 蜂鸣器），无网络时仍可判读", "S", "断网走查"],
       ["FR-10", "记录报警事件（含触发指标与时间戳）并可导出", "S", "导出并核对"],
       ["FR-11", "支持按键重新学习基线与静音", "C", "现场操作验证"]],
      [Cm(1.6), Cm(9.4), Cm(1.6), Cm(3.4)], size=8.6,
      align_map=["c", "l", "c", "l"])
table("表 28　非功能性需求", ["编号", "类别", "需求描述"],
      [["NFR-01", "安全", "全部电路 ≤ 12 V 直流，不接触市电；转动部件加防护罩"],
       ["NFR-02", "隐私", "不使用摄像头或麦克风，不采集任何可识别个人的数据"],
       ["NFR-03", "可靠性", "连续运行 ≥ 72 h 无重启；看门狗、掉电恢复、断网自愈"],
       ["NFR-04", "可维护性", "传感器可现场更换且免重新编译；基线可一键重学"],
       ["NFR-05", "成本", "单节点 ≤ ¥250（含耗材与备件，见表 7）"],
       ["NFR-06", "可访问性", "仪表板中英双语；状态用文字 + 图标 + 颜色三重编码；字号可调"],
       ["NFR-07", "互操作性", "数据模型带版本号并有文档，可被第三方平台订阅复用"],
       ["NFR-08", "环境", "外壳防护 ≥ IP54，工作温度 0 ~ 50 ℃"]],
      [Cm(1.8), Cm(1.9), Cm(12.3)], size=8.6, align_map=["c", "c", "l"])
table("表 29　验收标准与验证方法", ["编号", "验收判据", "目标值", "验证方法"],
      [["AC-01", "振动采样率准确度", "误差 ≤ 0.5 %", "连续采集 1 s 统计实际点数"],
       ["AC-02", "频谱频率准确度", "误差 ≤ 1 % 或 ≤ Δf", "音箱振膜 + 100 Hz 正弦"],
       ["AC-03", "幅值准确度（比对）", "偏差 ≤ 10 %", "与手持测振仪同点同时比对"],
       ["AC-04", "温度测量准确度", "± 0.5 ℃", "冰水浴与温水比对"],
       ["AC-05", "工况变化识别", "3 类中至少 2 类在 60 s 内检出并分级正确",
        "不平衡 / 松动 / 负载变化对照实验"],
       ["AC-06", "误报控制", "正常工况 2 h 内假阳性 ≤ 1 次", "正常工况长时运行"],
       ["AC-07", "报警端到端延迟", "≤ 5 s", "注入越限信号并计时"],
       ["AC-08", "连续运行能力", "≥ 72 h 无重启，丢数率 ≤ 1 %", "不间断运行并统计记录"],
       ["AC-09", "断网恢复", "≤ 30 s 内自动重连并补传", "拔插网线测试"],
       ["AC-10", "仪表板刷新与历史", "刷新 ≤ 5 s；历史可查 ≥ 4 周", "走查 + 数据库查询"],
       ["AC-11", "单节点成本", "≤ ¥250", "按第二章清单核算"],
       ["AC-12", "重复性", "变异系数 ≤ 3 %", "同一工况重复测量 10 次"]],
      [Cm(1.6), Cm(4.8), Cm(5.0), Cm(4.6)], size=8.6,
      align_map=["c", "l", "l", "l"])

# ================================================================= 十
h("十、与手册要求的逐条对照", 1)
table("表 30　手册要求与本方案的对应关系", ["手册要求", "本方案如何满足", "证据"],
      [["明确定义的客户需求、可衡量的验收标准",
        "见第九章表 27 ~ 表 29", "第九章"],
       ["可工作的集成原型", "ESP32 单板完成采集 → 处理 → 通信 → 决策全链路",
        "步骤 9 的端到端演示"],
       ["至少两个有意义的数据输入并做校准", "振动与温度两路，均可独立校准",
        "校准报告（重力翻转 / 水浴 / 比对）"],
       ["有意义的本地处理与决策逻辑",
        "本地完成 2048 点 FFT、特征提取、基线自学习与三级判别", "固件源码 + 单元测试记录"],
       ["适当的网络通信（Wi-Fi / MQTT）", "ESP32 内置 Wi-Fi + MQTT，10 s 周期 + 事件即时",
        "抓包记录 + 端到端延迟测试"],
       ["仪表板：实时与历史视图、报警与建议机制",
        "Node-RED dashboard：状态卡 + 24 h 趋势 + 频谱 + 分级报警与建议；SQLite 存历史",
        "界面走查 + 非项目成员可用性测试"],
       ["单元 / 集成 / 系统测试与实验证据",
        "第九章至第十一章的校准与实验流程", "台账「6-测试记录」「7-工况实验记录」"],
       ["专业实践：版本控制、会议记录、问题跟踪、风险登记册、受控文档",
        "Git 仓库 + 受控文档台账（需求与验收、物料清单、风险登记册、责任矩阵、十二周计划）",
        "Git 仓库 + Excel 台账"],
       ["负责任设计：隐私、安全、伦理、可访问性",
        "无摄像头、无个人数据；全程 ≤ 12 V；仪表板中英双语、状态三重编码",
        "伦理 / 隐私 / 安全筛查表"],
       ["跨国部署：讨论在 UL 与 LICHPU 复制的可能性",
        "器件在两地均可采购；外壳防护、温度阈值与界面语言按环境调整", "见第十二章"]],
      [Cm(5.4), Cm(6.6), Cm(3.8)], size=8.2, align_map=["l", "l", "l"])

# ================================================================= 十一
h("十一、取舍与局限", 1)
para("手册的“优秀”评价档明确要求“诚实地陈述局限性”。"
     "以下内容应当直接写入最终报告的“讨论、局限性与经验教训”一节。", indent=True)
table("表 31　局限与应对", ["局限", "影响", "缓解措施与报告中的表述"],
      [["无转速传感器，不能做阶次归一化",
        "转速波动较大时判据会退化，仅在额定转速附近可靠",
        "台架风机转速稳定；报告中写为“假定恒转速工况下的变化识别”，"
        "并用转速变化实验量化该假设的边界"],
       ["无包络解调，对轴承早期点蚀不敏感",
        "无法在缺陷极早期发现轴承问题",
        "定位为识别不平衡、积尘、松动、负载变化四类——恰好是校园风机最常见故障；"
        "轴承诊断列为未来工作"],
       ["ESP32 无硬件 RTC",
        "断网时时间戳退化为“运行时长”",
        "联网时用 NTP 同步；补传时以服务器接收时间标记；可选加装 DS3231（约 ¥8）"],
       ["本地无显示屏", "现场判读需连接手机或电脑",
        "演示时投大屏更直观；RGB LED + 蜂鸣器提供无网络时的最低限度指示"],
       ["Node-RED 依赖笔记本在线", "笔记本不在线时看不到实时界面",
        "节点侧 microSD 独立留存历史，补传后数据不丢；历史要求不依赖服务端在线"]],
      [Cm(3.4), Cm(4.4), Cm(8.0)], size=8.4, align_map=["l", "l", "l"])

# ================================================================= 十二
h("十二、跨国复制与部署要点", 1)
para("手册要求讨论在 UL 与 LICHPU 两地复制的可能性，包括采购、服务可用性、语言与环境差异。"
     "本方案从器件、供电、环境、网络、设备与法规六个维度做了对照。", indent=True)
table("表 32　UL 与 LICHPU 复制差异与对策", ["维度", "差异", "对策"],
      [["采购与成本", "中国采购约 ¥211；爱尔兰本地分销约 €90 ~ €110",
        "集中采购后随身携带（注意锂电池航空限制）；批量采购摊薄运费与关税"],
       ["供电", "爱尔兰 230 V / 中国 220 V，均为 50 Hz",
        "统一使用 12 V 适配器，并分别配 CE 与 CCC 认证型号；节点侧接口不变"],
       ["气候与环境", "爱尔兰湿冷多雨易凝露；中国夏季高温高湿",
        "IP54 外壳 + 透气防水阀；接线端子防潮；温度阈值用与环境温差而非绝对温度"],
       ["网络", "两地校园 Wi-Fi 均可用 2.4 GHz 信道 1 ~ 13",
        "节点不硬编码信道；MQTT Broker 在两地各部署或使用云端；必要时降级为本地记录"],
       ["设备与基线", "两地风机品牌、叶片数与转速不同",
        "基线自学习使阈值自适应，无需按品牌重新标定；仅需一次现场基线采集"],
       ["语言与可访问性", "两地语言不同",
        "仪表板中英双语；状态用文字 + 图标 + 颜色三重编码"],
       ["法规与数据保护", "爱尔兰适用 GDPR；中国适用 PIPL",
        "不采集个人数据，仅聚合设备物理量，天然满足数据最小化；数据限域存储、保留期 ≤ 4 周"]],
      [Cm(2.4), Cm(5.8), Cm(7.4)], size=8.6, align_map=["l", "l", "l"])

# ================================================================= 十三
h("十三、后续扩展方向", 1)
para("若中期评审后时间仍有富余，可按下列顺序逐步增强；每一步都是独立可交付的增量，"
     "不会推翻已有工作。", indent=True)
table("表 33　可选增强项（时间允许时按顺序实施）", ["顺序", "增强项", "增加成本", "增加工作量"],
      [["1", "加装光电 / 霍尔转速传感器，用中断 + micros() 测周期", "¥5", "1 ~ 2 天"],
       ["2", "加入转速分档基线与阶次谱", "¥0", "2 ~ 3 天"],
       ["3", "加入包络解调（可在笔记本侧用 Python 离线做）", "¥0", "2 ~ 3 天"],
       ["4", "加装 OLED 本地界面（可选）", "¥12", "1 ~ 2 天"],
       ["5", "服务端由 Node-RED 升级为 InfluxDB + Grafana", "¥0", "1 ~ 2 天"],
       ["6", "部署到校园真实风机，进行 4 周试运行", "¥0（复用现有节点）", "持续运行 + 每周巡检"]],
      [Cm(1.3), Cm(8.9), Cm(2.9), Cm(3.0)], size=8.6,
      align_map=["c", "l", "c", "c"])

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(12)
setfont(p.add_run("（实施方案完）"), size=10, ea="黑体",
        color=RGBColor(0x70, 0x70, 0x70))
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.save(OUT)
print("saved:", OUT)
