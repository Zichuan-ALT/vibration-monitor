# -*- coding: utf-8 -*-
"""生成《项目 8.1 振动状态监测器 — 实施方案》（免焊接 · 集成模块版，预算 ≤ ¥150）。"""
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
setfont(p.add_run("实施方案"), size=15, bold=True, ea="黑体", color=ACCENT)

table("", ["项目", "内容"],
      [["文件性质", "本项目单节点样机的完整实施方案，覆盖选型、耗材、接线、施工、"
                "固件、服务端、校准与验证"],
       ["技术路线", "ESP32 单板 + MPU6050（I²C，加速度 + 板载温度）+ DS18B20 表面温度 "
                "+ Node-RED 仪表板"],
       ["装配方式", "**全程免焊接**：模块均选购已焊排针版，杜邦线直插，"
                "电源经面包板电源轨分配；不需要电烙铁，也不需要免焊端子"],
       ["硬件成本", "**实际支出 ¥138.85**（含备用 ESP32）；预算上限 ¥150，结余 ¥11.15"],
       ["实施工时", "约 11 ~ 15 个工作日（免焊接后接线与调试时间明显缩短）"],
       ["核心内容", "13 个施工步骤，每步给出验收标准与常见问题；"
                "11 项需求与 13 条验收标准；按实际采购的物料清单与成本核算"],
       ["版本 / 日期", f"V3.2（按实际采购器件更新物料与成本）　{TODAY.strftime('%Y-%m-%d')}"]],
      [Cm(2.8), Cm(13.6)], size=9.5, align_map=["c", "l"])

# ================================================================= 一
h("一、设计思路与总览", 1)
h("1.1 设计思路", 2)
para("本项目受两条约束：**硬件预算不超过 ¥150**，以及**尽量减少焊接操作**"
     "（实验室焊接条件有限，且焊接质量会直接影响调试效率与可靠性）。"
     "两条约束指向同一个解法：用**集成度更高的模块 + 杜邦线直插**，"
     "把元件数量和手工工序同时压下来。", indent=True)
para("同时，评估权重中技术设计与实现（C）占 25%，而测试、验证与数据分析（D）占 20%、"
     "报告（E）占 15%、专业执行（B）占 15%。四分之三的分数来自“系统做完、测过、"
     "记录清楚”，因此本方案把省下的时间全部投入到校准、对照实验与数据分析上。",
     indent=True)

h("1.2 关键设计选择", 2)
table("表 1　关键设计选择与理由", ["设计环节", "本方案的选择", "理由"],
      [["主控", "ESP32-WROOM-32 单板（内置 Wi-Fi，已焊排针版）",
        "一块板同时完成采集、计算与联网，省掉模块间通信与额外调试；排针已焊，开箱即用"],
       ["振动与温度采集", "MPU6050 模块（GY-521，I²C）",
        "一颗芯片同时提供三轴加速度、三轴陀螺与板载温度；只需 4 根线，"
        "且省掉 ADXL345 与独立温度芯片"],
       ["表面温度", "DS18B20 防水探头（1-Wire）",
        "补充一路可直接贴附在风机外壳上的温度，比芯片结温更贴近实际工况"],
       ["本地指示", "板载 LED + 板载 BOOT 按键 + 有源蜂鸣器模块",
        "开发板本身自带 LED 与按键，不必另购，也无需焊接限流电阻与三极管"],
       ["装配方式", "全部杜邦线直插 + 免焊螺丝端子",
        "**零焊接工序**，出问题可立即换线换模块，调试速度显著提升"],
       ["服务端", "Node-RED 单进程 + SQLite（跑在笔记本上）",
        "一个进程同时充当 MQTT Broker、历史库与仪表板，零硬件采购"],
       ["供电", "ESP32 用 USB 充电器；5 V 风扇用另一个充电器经 USB 转杜邦线供电",
        "两路分开，避免电机噪声串入数字电源；都用手边现有的 USB 电源，成本为零"],
       ["硬件成本", "¥94（含备用板 ¥119）", "见第二章逐项清单，留有 ¥31 以上余量"],
       ["实施工时", "约 11 ~ 15 个工作日", "免焊接后接线与排故时间缩短约 2 天"]],
      [Cm(2.4), Cm(6.0), Cm(8.0)], size=8.5, align_map=["l", "l", "l"])

h("1.3 系统总览", 2)
figure("figS1_arch.png", "图 S1　系统架构", 15.8)
rich([("关键点：", True),
      ("MPU6050 与 DS18B20 都是数字接口（I²C 与 1-Wire），因此整个系统不需要 ADC、"
       "不需要模拟前端、也不需要信号调理电路。板载 LED 与按键省掉两个外设及其限流电阻，"
       "有源蜂鸣器模块自带驱动省掉三极管——这些都直接减少了焊接点。", False)])

h("1.4 总览", 2)
table("表 2　总览", ["项目", "内容"],
      [["硬件成本", "¥94（主要器件 ¥73 + 连接件 ¥8 + 结构与电源 ¥13）；"
                "含备用 ESP32 板 ¥119。预算上限 ¥150，余量 ≥ ¥31"],
       ["服务端成本", "¥0（使用现有笔记本 + Node-RED + SQLite，不采购服务器）"],
       ["装配工序", "零焊接：所有连接均为杜邦线直插或免焊螺丝端子"],
       ["实施工时", "约 11 ~ 15 个工作日（不含报告、海报与答辩准备）"],
       ["里程碑", "第 5 ~ 6 周即可得到可演示的完整系统（采集 → 判别 → 上报 → 仪表板）"],
       ["交付物", "节点样机 1 台、固件源码 1 套、Node-RED 流程 1 套、"
               "实验数据集 1 份、校准与测试记录 1 套"]],
      [Cm(2.6), Cm(13.8)], size=8.8, align_map=["c", "l"])

# ================================================================= 二
h("二、物料清单（实际采购）", 1)
para("本章按“主要器件 / 连接件 / 结构与电源”三类列出**实际采购**的型号与价格。"
     "全部为**已焊排针**的成品模块；传感器与主控之间用杜邦线直插，"
     "电源经面包板电源轨分配——**装配过程不需要电烙铁，也不需要免焊端子和 12 V 升压线**。",
     indent=True)

h("2.1 主要器件", 2)
table("表 3　主要器件清单（实际采购）", ["#", "物料", "实际采购型号 / 规格", "数量", "单价 ¥", "小计 ¥"],
      [["1", "主控板", "ESP32-WROOM-32 开发板（**CH340C，已焊针 30Pin，Type-C**）",
        "2", "23.00", "46.00"],
       ["2", "振动 + 温度", "MPU-6050 三轴重力加速度模块（GY-521，已焊针，I²C）",
        "1", "13.50", "13.50"],
       ["3", "表面温度", "DS18B20 带防水探头转接模块 + 传感器 + 杜邦线", "1", "8.50", "8.50"],
       ["4", "存储卡座", "Micro SD 卡模块 SPI 接口（已焊针，板载电平转换）",
        "1", "2.30", "2.30"],
       ["5", "存储卡", "microSD 32 GB Class10 A1（联想，全新带 5 年换新）",
        "1", "35.90", "35.90"],
       ["6", "声音指示", "有源蜂鸣器模块（3.3 ~ 5 V，**低电平触发**）", "1", "1.50", "1.50"],
       ["7", "演示设备", "静音小风扇 5 V 9225（92 mm，**1700 RPM**）", "1", "4.80", "4.80"],
       ["", "小计", "含 1 块备用 ESP32 板", "", "", "112.50"]],
      [Cm(0.8), Cm(2.4), Cm(7.8), Cm(1.2), Cm(1.6), Cm(1.6)], size=8.5,
      align_map=["c", "l", "l", "c", "c", "c"])

h("2.2 连接件", 2)
table("表 4　连接件清单（实际采购）", ["#", "物料", "规格 / 说明", "数量", "单价 ¥", "小计 ¥"],
      [["8", "数据线", "TYPE-C 数据线 1 m（**可传输数据**，非纯充电线）", "1", "2.20", "2.20"],
       ["9", "风扇供电", "USB 公头转杜邦 2.54 mm 线（取 USB 5 V 给风扇）", "1", "4.80", "4.80"],
       ["10", "杜邦线", "母-母 / 母-公 20 cm（信号线 + 电源线）", "1 套", "**学校提供**", "0"],
       ["11", "面包板", "400 孔（3V3 / GND / 5V 电源轨分配）", "1", "学校 / 自有", "0"],
       ["", "小计",
        "**不需要洞洞板、排针、焊锡、助焊剂、热缩管，也不需要免焊端子与 12 V 升压线**",
        "", "", "7.00"]],
      [Cm(0.8), Cm(2.4), Cm(7.8), Cm(1.2), Cm(1.6), Cm(1.6)], size=8.5,
      align_map=["c", "l", "l", "c", "c", "c"])

h("2.3 结构与电源", 2)
table("表 5　结构与电源清单", ["#", "物料", "规格 / 说明", "数量", "单价 ¥", "小计 ¥"],
      [["12", "外壳与底板", "快递纸盒 + 木条 200×100 mm（自备）", "1", "自备", "0"],
       ["13", "固定与配重", "双面胶、尼龙扎带、硬币 / 螺母（自备）", "1 套", "自备", "0"],
       ["14", "电源", "USB 充电器 ×2（ESP32 与风扇各一，现有）", "2", "现有", "0"],
       ["", "小计", "电源使用手边现有 USB 充电器，不计入采购", "", "", "0"]],
      [Cm(0.8), Cm(2.4), Cm(7.8), Cm(1.2), Cm(1.6), Cm(1.6)], size=8.5,
      align_map=["c", "l", "l", "c", "c", "c"])

h("2.4 需要准备的工具", 2)
para("**本方案不需要电烙铁，也不需要免焊端子。** 只需以下几样，实验室通常都能借到：", indent=True)
bullets([
    "万用表 —— 每一步接线后先量通断与电压，再上电；",
    "螺丝刀、斜口钳（剪扎带、剥电源线）；",
    "**USB-A 转 Type-C 数据线**（已采购 1 m）—— 必须是数据线，纯充电线电脑识别不到板子；",
    "**CH341SER 驱动** —— 本方案采购的板是 CH340，装 WCH 官方的 CH341SER（见步骤 1）；",
    "手持测温计（可选）—— 校准 DS18B20 用；",
    "手持测振仪 —— 幅值比对用，**必须先确认实验室能否借到**（风险 R-02 的前提）。",
], size=10)

h("2.5 到货核对要点", 2)
table("表 6　到货核对与替代型号", ["物料", "本次采购", "到货先做", "可用替代"],
      [["主控板", "ESP32-WROOM-32 · CH340C · 已焊针 30Pin · Type-C",
        "验货确认**排针已焊**；装 CH341SER 后设备管理器出现 COM 口",
        "ESP32-S3 / C3（引脚需重排）"],
       ["振动", "MPU-6050 三轴加速度模块（已焊针）",
        "读 WHO_AM_I 应为 **0x68**；静止 Z 轴约 1.000 g，零偏 ≤ ±0.2 g",
        "ICM-42688-P、ADXL345（需改驱动）"],
       ["表面温度", "DS18B20 带防水探头转接模块",
        "探头裸线接**模块自带的绿色端子**；读数不应是 85 ℃ 或 −127 ℃",
        "SHT31（I²C）"],
       ["存储", "Micro SD 卡模块（SPI + 电平转换）",
        "**VCC 先试 3.3 V**，识别不到再改 5 V；接前量 CS/MOSI/SCK 对 GND",
        "直接用 ESP32 片内 SPIFFS（容量小）"],
       ["存储卡", "microSD 16 / 32 GB Class10 A1（全新品牌）",
        "**拷 4 GB 文件验证容量**（二手/扩容卡会循环覆盖）", "—"],
       ["声音指示", "有源蜂鸣器模块（**低电平触发**）",
        "确认有源且带驱动；**GPIO25 输出 LOW 时响**，代码需反相", "高电平触发版（更省心）"],
       ["演示风扇", "静音风扇 5 V 9225（1700 RPM，XH2.54-2P）",
        "USB 公头转杜邦线供电：**红→5V、黑→GND**，白绿悬空；杜邦母头直接插 XH 公针",
        "任意 120 mm 风扇（转速 1000 ~ 2000 r/min）"]],
      [Cm(1.8), Cm(4.0), Cm(5.4), Cm(3.4)], size=8.2,
      align_map=["l", "l", "l", "l"])

h("2.6 成本汇总", 2)
table("表 7　成本汇总（预算上限 ¥150）", ["项目", "金额 ¥", "说明"],
      [["主要器件（表 3）", "112.50",
        "ESP32 ×2、MPU6050、DS18B20、microSD 卡座与卡、蜂鸣器、演示风扇"],
       ["连接件（表 4）", "7.00", "数据线、USB 转杜邦线；杜邦线与面包板由学校提供"],
       ["结构与电源（表 5）", "0", "外壳、底板、配重自备；电源用现有 USB 充电器"],
       ["实际支出合计", "119.50", "4 人分摊约 ¥30 / 人"],
       ["预算上限（老师要求）", "150", "—"],
       ["结余", "30.50", "**达标，余量 ¥30.50**（约 20 %）"],
       ["按预算单价口径（对照）", "约 135", "编制方案时的估价，实际支出低于估价"]],
      [Cm(4.4), Cm(2.6), Cm(9.0)], size=8.5, align_map=["l", "c", "l"])
rich([("与初始方案的差异说明：", True),
      ("初始方案曾列入洞洞板、排针、焊锡、助焊剂、热缩管（约 ¥40）、KF301 免焊端子（约 ¥13）、"
       "12 V 机箱风扇与升压线（约 ¥19）。实际采购中全部取消或替换："
       "① **已焊排针模块 + 杜邦线直插**（杜邦线由学校提供），不需要焊接；"
       "② 全项目唯一的裸线（DS18B20 探头引线）由**该模块自带的绿色端子**解决，不需要免焊端子；"
       "③ 改用 **5 V 9225 风扇**（1700 RPM），用 USB 公头转杜邦线直接供 5 V，**不需要 12 V 升压线**。"
       "因此实际支出由估价 ¥135 降至 **¥119.50**，结余 ¥30.50（约 20 %）。", False)])

# ================================================================= 三
h("三、接线与装配", 1)
h("3.1 接线图", 2)
figure("figS_wiring.png", "图 S2　接线示意图（全程免焊接，④ 为声音指示）", 15.8)

h("3.2 引脚分配", 2)
table("表 8　ESP32 引脚分配", ["引脚", "功能", "连接对象", "说明"],
      [["GPIO21", "I²C SDA", "MPU6050 SDA", "I²C 数据线，400 kHz"],
       ["GPIO22", "I²C SCL", "MPU6050 SCL", "I²C 时钟线，400 kHz"],
       ["GPIO26", "外部中断", "MPU6050 INT", "数据就绪 / FIFO 水位，可选"],
       ["GPIO5", "1-Wire", "DS18B20 数据", "4.7 kΩ 上拉到 3V3（模块自带）"],
       ["GPIO18", "SPI SCK", "microSD 时钟", "SPI 总线时钟"],
       ["GPIO19", "SPI MISO", "microSD 数据输入", "SPI 总线 MISO"],
       ["GPIO23", "SPI MOSI", "microSD 数据输出", "SPI 总线 MOSI"],
       ["GPIO13", "SPI CS", "microSD 片选", "片选，软件控制"],
       ["GPIO25", "数字输出", "有源蜂鸣器模块信号脚",
        "模块 3 线：VCC→3V3、GND、信号→GPIO25；**本模块低电平触发，输出 LOW 时响**，"
        "上电前需先拉高（见 5.3 ⑤）"],
       ["GPIO2", "板载 LED", "**无需接线**", "开发板自带，用闪烁模式表示状态"],
       ["GPIO0", "板载按键", "**无需接线**", "开发板自带 BOOT 键，用作功能键"],
       ["GPIO1 / 3", "USB 串口", "调试与烧录", "板载 USB 完成"],
       ["3V3 / GND", "电源", "MPU6050、DS18B20、蜂鸣器", "**一律 3.3 V，禁止接 5 V**"],
       ["5V / VIN", "电源", "microSD 卡座模块",
        "**该模块 VCC 要求 4.5 ~ 5.5 V，必须接 5 V**（板载稳压再给卡供 3.3 V）"]],
      [Cm(2.4), Cm(2.8), Cm(4.2), Cm(7.0)], size=8.5,
      align_map=["c", "l", "l", "l"])

h("3.3 供电与接地", 2)
bullets([
    "**电源分配**：ESP32 的 3V3 / GND / 5V 用**母-公杜邦线**接到面包板电源轨，"
    "各模块的 VCC / GND 从电源轨取电——ESP32 上 3V3 只有 1 个针脚，"
    "要供 3 个模块，直接叠插在针脚上不可靠，面包板是免焊、直观的解法。",
    "ESP32 由 USB 充电器（5 V）供电；**5 V 9225 风扇由另一个充电器或充电宝经 "
    "USB 公头转杜邦线供电**（红线→5 V、黑线→GND，另两根悬空）。"
    "**两路分开**，避免电机噪声串入数字电源——这是本项目最容易忽视又最容易导致"
    "数据异常的问题。",
    "所有模块的 GND 与 ESP32 共地。MPU6050、DS18B20、蜂鸣器接 3V3；"
    "**microSD 卡座模块要求 VCC 4.5 ~ 5.5 V，接 ESP32 的 5V / VIN**。",
    "信号线尽量短（≤ 20 cm），并远离风扇电源线。",
    "USB 充电器与充电宝属于“手边现有”，不计入采购预算。",
])

h("3.4 装配顺序（无焊接工序）", 2)
table("表 9　装配顺序与注意事项", ["顺序", "动作", "注意事项"],
      [["1", "检查各模块排针是否已经焊好", "**这是唯一可能翻车的地方**："
        "若收到的是未焊排针的散件，请换货或请实验室代焊，不要自己硬上"],
       ["2", "用万用表确认开发板 3V3 与 GND 输出电压正常", "先只上电不接任何模块"],
       ["3", "用杜邦线把 MPU6050 接到 GPIO21/22 与 3V3/GND",
        "I²C 只有 4 根线；接完先跑读数自检"],
       ["4", "DS18B20 探头三根裸线拧进免焊螺丝端子，再插杜邦线到 GPIO5",
        "红线 VCC、黑线 GND、黄线数据；不要接反"],
       ["5", "用杜邦线接 microSD 模块（GPIO18/19/23/13）", "每接一路都要重新做一次自检"],
       ["6", "蜂鸣器模块（3 线）：VCC 接 3V3、GND，信号脚接 GPIO25",
        "模块自带驱动无需三极管；**低电平触发**，先烧录固件再上电可避免长鸣"],
       ["7", "把模块与走线放进塑料盒，用双面胶和扎带固定",
        "DS18B20 探头要伸出盒外，便于贴到风机上"],
       ["8", "把风扇固定在木板上，MPU6050 用双面胶或螺栓贴在风扇外壳",
        "安装要尽量**刚性**：贴紧外壳、接触面积大，避免用厚海绵垫"]],
      [Cm(1.2), Cm(7.6), Cm(7.6)], size=8.5, align_map=["c", "l", "l"])

# ================================================================= 四
h("四、分步实施（13 步）", 1)
para("每一步都遵循同一个原则：**只增加一个变量，做完立即验证并记录**。"
     "免焊接方案下，步骤 1 ~ 4 的总耗时比上一版缩短约 1 天。", indent=True)

steps = [
    ("表 10　步骤 0：采购与工具准备（0.5 天）",
     "备齐第二章全部物料，安装开发与服务端软件",
     "① 按表 3 ~ 表 5 下单（**全部选已焊排针/带驱动模块**，ESP32 买两块）；"
     "② 安装 Arduino IDE 2.x 并在板管理器中添加 ESP32 支持包；"
     "③ 安装 Node.js LTS，执行 npm install -g --unsafe-perm node-red；"
     "④ 在 Node-RED 的“节点管理”中安装 node-red-dashboard、node-red-node-sqlite、"
     "node-red-contrib-aedes 三个模块",
     "Arduino IDE 能编译 ESP32 的 Blink 示例；浏览器打开 http://localhost:1880 看到编辑器；"
     "三个 palette 出现在左侧节点栏",
     "aedes 模块负责充当 MQTT Broker，装了它就不必再单独安装 Mosquitto。"
     "**下单时务必确认勾选“已焊排针”**，否则会收到一堆要自己焊的散件。"
     "到货当天先验货：排针是否已焊、随附线缆是不是数据线"),
    ("表 11　步骤 1：点亮最小系统（0.5 天）",
     "确认工具链、驱动、串口全部正常",
     "① 按板载 USB 转串口芯片型号安装驱动（**本方案采购的板是 CH340 → 装 WCH 的 CH341SER 驱动**；"
     "其他板可能是 CP2102 → Silicon Labs CP210x VCP，或 CH9102 → WCH）；"
     "② 用 USB-A 转 Type-C **数据线**接入电脑，在设备管理器确认出现串口；"
     "③ 烧录 Blink + 串口打印，串口监视器波特率 115200",
     "设备管理器识别到 COM 口；板载 LED 闪烁；串口能看到周期输出",
     "这一步要排除三个坑，按现象分头排查："
     "① 设备管理器出现黄色感叹号 → 驱动没装或装错型号，按芯片型号重装；"
     "② 插上完全没反应、连未知设备都没有 → 多半是线的问题，换一根 USB-A 转 Type-C "
     "**数据线**（C-to-C 线或纯充电线常常不传数据）；③ 上传报 Failed to connect → "
     "按住 BOOT 键再点上传。**先换线，再怀疑板子。**"),
    ("表 12　步骤 2：接通 MPU6050（1 天）",
     "把加速度计读通，拿到可信的加速度与板载温度原始值",
     "按表 8 接线（SDA→GPIO21、SCL→GPIO22）→ 先读 WHO_AM_I 寄存器（0x75，"
     "应返回 0x68）→ 再读三轴加速度与板载温度",
     "WHO_AM_I 返回 0x68；静止时某一轴约为 ±1 g；轻敲桌面数值明显跳变；"
     "板载温度读数接近室温（略高）",
     "**本模块为国产芯片**，功能正常但批次差异较大，到货必须做三步自检：读 WHO_AM_I 应为 0x68、"
     "静止 Z 轴约 1.000 g（零偏 ≤ ±0.2 g）、连采 1000 点看噪声量级；任一项不过立即换货。"
     "I²C 地址是 0x68（AD0 接地）或 0x69（AD0 接高）。"
     "若扫描不到设备，先用 I²C 扫描程序确认地址；"
     "量程建议设为 ±4 g（8192 LSB/g，兼顾分辨率与量程）"),
    ("表 13　步骤 3：接通 DS18B20（0.5 天）",
     "读到可直接贴附外壳的表面温度",
     "探头裸线拧进免焊螺丝端子，信号端接 GPIO5；用 OneWire + DallasTemperature 库读取",
     "室温读数合理；手指捏住探头，温度在数秒内上升 1 ~ 2 ℃",
     "若读到 −127 ℃，说明时序或上拉有问题；先只接一个 DS18B20，"
     "多探头需要先搜索 ROM 地址。**不要接反电源线**，反接会烧探头"),
    ("表 14　步骤 4：接通 microSD 存储（0.5 天）",
     "能写能读，为断网缓存做准备",
     "按表 8 接线（SCK→18、MISO→19、MOSI→23、CS→13；VCC 接 5V / VIN）；接前先量信号脚电压；用 SD 库挂载并写一个 CSV "
     "文件再读回",
     "文件创建成功，内容读回一致",
     "**本模块 VCC 要求 4.5 ~ 5.5 V，接 3V3 会识别不到卡**；文件名用 8.3 短名兼容性最好；"
     "挂载失败先把 SPI 时钟降到 4 MHz 试试。CS 用 GPIO13 而不是 GPIO2，"
     "因为 GPIO2 是板载 LED"),
    ("表 15　步骤 5：完成定时采样与 FFT（2 ~ 3 天）",
     "把时域波形变成可信的频谱——这是全项目的技术核心",
     "① 用微秒级定时循环实现 1 kSPS × 1024 点采样；② 去均值 + Hanning 窗；"
     "③ 用 arduinoFFT 做 1024 点 FFT；④ 换算幅值谱",
     "采 1 s 统计点数为 1000 ± 5；把传感器贴在小音箱振膜上放 100 Hz 正弦，"
     "谱峰出现在 100 Hz ± 1 %；贴在风机外壳上时，基频带内应出现约 25 Hz 的峰",
     "**不要用 delay() 循环采样**，采样率会飘、频谱全假；"
     "把采样放在高优先级任务并绑定到 ESP32 的 core 1，网络留在 core 0，"
     "否则 Wi-Fi 会打断采样节拍。MPU6050 建议用 FIFO 突发读取，"
     "避免每个采样点单独发起一次 I²C 事务"),
    ("表 16　步骤 6：实现特征计算（1 天）",
     "把频谱与时域数据压缩成几个可判读的指标",
     "计算 RMS、峰值、速度有效值（10 ~ 250 Hz 积分）以及基频带（15 ~ 35 Hz）"
     "内的最大幅值及其频率；同时记录 DS18B20 表面温度与 MPU6050 板载温度",
     "各指标随时间平稳；用手挡风道时 RMS 与温度同时上升",
     "没有转速传感器时，**不要去找精确的 1×**，改用“基频带内最大峰值”作为指标，"
     "这样转速小范围波动时结果依然稳定"),
    ("表 17　步骤 7：基线自学习与三级判别（1 ~ 2 天）",
     "让节点自己判断设备状态",
     "① 开机连续采集 60 s，算振动 RMS 的 μ 与 σ；② 按 μ+3σ / μ+6σ 与温度增量"
     "+8 ℃ / +15 ℃ 分级；③ 加迟滞（连续 3 帧确认报警、连续 5 帧确认恢复）；"
     "④ 驱动板载 LED（闪烁模式）与蜂鸣器（**低电平触发：输出 LOW 时响**）。"
     "用 BUZZER_ON() / BUZZER_OFF() 两个宏隔离极性，并在 setup() 开头按"
     "\u201cINPUT_PULLUP → 输出置高 → 切 OUTPUT\u201d的顺序初始化，避免复位期间长鸣；"
     "若仍长鸣，在 I/O 与 3V3 之间并接 10 k\u03a9 上拉电阻",
     "叶轮贴配重后状态由正常升到注意或报警，撕掉后能自动恢复；"
     "正常工况连续运行 2 h 内假阳性不超过 1 次",
     "基线只在“连续 30 min 处于正常状态”时缓慢更新，"
     "否则会把已经劣化的状态学成新的正常，导致漏报。"
     "板载 LED 是单色，用闪烁频率区分三级状态（常亮 / 1 Hz / 4 Hz）"),
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
     "SQLite 节点要显式建表（见 6.2），否则写入会静默失败；"
     "ui_chart 默认只保留内存中的点数，长趋势要靠 SQLite 查询回填"),
    ("表 20　步骤 10：完成校准（1 天）",
     "让测量结果可追溯、可解释",
     "用重力翻转法标定加速度灵敏度与零偏；冰水浴与温水比对 DS18B20；"
     "与手持测振仪同点同时比对幅值；板载温度与 DS18B20 做同步记录以给出偏置",
     "见第七章表 25 的判据；全部记录填入台账「6-测试记录」",
     "没有振动台时，报告里必须写明这是“比对”而不是“绝对溯源”，并给出不确定度预算。"
     "**先确认实验室能借到手持测振仪与手持测温计**，否则相应判据要降级为相对趋势验证"),
    ("表 21　步骤 11：搭建演示台架并制造工况（1 天）",
     "能在台架上稳定复现三类工况变化",
     "把 USB 风扇固定到底板；准备 2 g / 5 g 配重片、可松动的底座螺丝、"
     "可遮挡 30 % / 60 % 的纸板",
     "每种工况施加方式可重复，施加前后特征值有明确差异",
     "配重用胶带贴在叶轮同一位置并记录质量；松螺丝要记录圈数；"
     "遮挡板要记录面积比例——**可复现性是评估组成 D 的核心**"),
    ("表 22　步骤 12：对照实验与数据分析（3 ~ 5 天）",
     "积累足以支撑结论的实验证据",
     "按表 26 的实验矩阵逐项执行，每种工况按规定次数重复，"
     "记录全部特征值与原始 1024 点数据",
     "每个工况的样本量达标；用箱线图与 ROC 曲线评估各判据的检出能力，"
     "并报告假阳性率",
     "数据当场记录，不要事后补；原始数据统一命名并存到 Git 仓库的 evidence 目录"),
]
for cap, goal, doing, accept, pit in steps:
    table(cap, ["项目", "内容"],
          [["目标", goal], ["做什么", doing], ["验收标准", accept], ["常见问题", pit]],
          [Cm(2.2), Cm(14.2)], size=8.5, align_map=["c", "l"])

# ================================================================= 五
h("五、固件实现要点", 1)
h("5.1 程序结构", 2)
figure("figS2_flow.png", "图 S3　固件主流程", 10.4)
bullets([
    "**双核分工**：采样任务用 xTaskCreatePinnedToCore 绑定到 core 1 并设高优先级，"
    "Wi-Fi 与 MQTT 留在 core 0。否则网络中断会打乱采样节拍，频谱失真。",
    "**依赖库**：MPU6050（或 i2cdevlib）、OneWire + DallasTemperature、arduinoFFT、"
    "PubSubClient、SD、ArduinoJson。全部为公共库，无需自己写底层驱动。",
    "**调试开关**：留一个串口命令（输入 r 重新学习基线、输入 d 打印一帧特征），"
    "现场调试效率会高很多。",
])

h("5.2 关键参数配置", 2)
table("表 23　固件关键参数", ["参数", "取值", "说明"],
      [["采样率 fs", "1000 Hz", "MPU6050 加速度计最高 ODR，与 DLPF 带宽联动"],
       ["帧长 N", "1024 点", "Δf = fs / N = 0.977 Hz"],
       ["帧周期", "约 1.02 s", "比上一版更短，报警响应更快"],
       ["分析带宽", "10 ~ 260 Hz", "受 MPU6050 内部低通（DLPF）限制"],
       ["量程", "± 4 g", "8192 LSB/g，兼顾分辨率与量程"],
       ["基频带", "15 ~ 35 Hz", "对应 900 ~ 2100 r/min；本方案风扇 1700 RPM ≈ **28.3 Hz**，正落在带内"],
       ["速度积分区间", "10 ~ 250 Hz", "超出台宽的分量剔除，避免积分发散"],
       ["基线自学习时长", "60 s", "开机后自动执行"],
       ["基线更新条件", "连续 30 min 正常", "防止把故障学成正常"],
       ["注意 / 报警阈值", "μ + 3σ / μ + 6σ", "振动 RMS"],
       ["温度注意 / 报警", "+8 ℃ / +15 ℃", "相对基线的增量（DS18B20）"],
       ["报警确认 / 恢复", "连续 3 帧 / 5 帧", "迟滞，抑制抖动"],
       ["蜂鸣器极性", "**LOW = 鸣响**",
        "本模块为**低电平触发**，与常见高电平模块相反；用 BUZZER_ON() / BUZZER_OFF() 宏隔离，换模块只改宏"],
       ["MQTT 上报周期", "10 s", "事件即时上报"]],
      [Cm(3.4), Cm(3.6), Cm(9.2)], size=8.5, align_map=["l", "c", "l"])

h("5.3 关键代码要点", 2)
rich([("① 定时采样（ESP32 双核，避开 Wi-Fi 干扰）", True)], indent=False,
     space_after=2)
code_block([
    "const int N = 1024;  const float FS = 1000.0;  const uint32_t T_US = 1000;",
    "void sampleTask(void *p) {",
    "  mpu.setFIFOEnabled(true);          // 用 FIFO 突发读，减少 I2C 事务",
    "  uint32_t next = micros();",
    "  for (int i = 0; i < N; i++) {",
    "    while ((int32_t)(micros() - next) < 0) { }   // 等到下一个采样时刻",
    "    next += T_US;",
    "    buf[i] = mpu.getAccelerationX();             // 读单轴（或从 FIFO 取）",
    "  }",
    "  vTaskDelete(NULL);",
    "}",
    "// setup() 中：采样任务绑到 core 1，优先级 3",
    "xTaskCreatePinnedToCore(sampleTask, \"sampling\", 4096, NULL, 3, NULL, 1);",
])
rich([("② FFT（arduinoFFT）", True)], indent=False, space_after=2)
code_block([
    "double vReal[N], vImag[N];",
    "double mean = 0; for (int i = 0; i < N; i++) mean += buf[i]; mean /= N;",
    "for (int i = 0; i < N; i++) { vReal[i] = buf[i] - mean; vImag[i] = 0; }",
    "FFT.Windowing(vReal, N, FFT_WIN_TYP_HANN, FFT_FORWARD);   // Hanning 窗",
    "FFT.Compute(vReal, vImag, N, FFT_FORWARD);",
    "FFT.ComplexToMagnitude(vReal, vImag, N);",
    "double fPeak = FFT.MajorPeak(vReal, N, FS);               // 主频",
])
rich([("③ 速度有效值（10 ~ 250 Hz 积分，单位 mm/s）", True)], indent=False,
     space_after=2)
code_block([
    "double s = 0;",
    "for (int k = 1; k < N / 2; k++) {",
    "  double f = k * FS / N;",
    "  if (f < 10 || f > 250) continue;",
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
rich([("⑤ 蜂鸣器与 LED 极性（本模块为低电平触发）", True)], indent=False, space_after=2)
code_block([
    "// 本模块 I/O 拉低时鸣响，与常见\u201c高电平触发\u201d模块相反",
    "#define BUZZER_PIN   25",
    "#define BUZZER_ON()  digitalWrite(BUZZER_PIN, LOW)    // 低电平触发",
    "#define BUZZER_OFF() digitalWrite(BUZZER_PIN, HIGH)",
    "",
    "void setup() {",
    "  // 无毛刺初始化：先内部上拉 → 输出寄存器置高 → 再切输出",
    "  // ESP32 复位期间 GPIO 为高阻态，低电平触发模块可能长鸣，",
    "  // 这样写能把误响窗口压到最短（毫秒级）",
    "  pinMode(BUZZER_PIN, INPUT_PULLUP);",
    "  digitalWrite(BUZZER_PIN, HIGH);",
    "  pinMode(BUZZER_PIN, OUTPUT);",
    "  // ... 其余初始化",
    "}",
    "",
    "// 报警：BUZZER_ON();   恢复正常：BUZZER_OFF();",
    "// 若上电仍长鸣：在 I/O 与 3V3 之间并接 10 k\u03a9 上拉电阻（实验室常备）",
])

# ================================================================= 六
h("六、服务端实现（Node-RED）", 1)
para("本章按**已实测跑通**的配置给出。Windows 下 Node-RED 的安装步骤与七个高频故障"
     "见附录 A。", indent=True)

h("6.1 组件与配置", 2)
table("表 24　Node-RED 组件与配置要点（实测）", ["顺序", "节点 / 组件", "配置要点"],
      [["1", "aedes broker", "端口 **1883**；仅监听局域网，勿暴露公网"],
       ["2", "mqtt-broker（配置节点）",
        "**指向 127.0.0.1:1883**。注意：`mqtt in` 必须引用这个配置节点，"
        "而不是 aedes broker 节点，否则永远收不到数据且无报错"],
       ["3", "mqtt in", "主题 `smartcampus/vib/+/state`，QoS 1，datatype 选 **JSON**"],
       ["4", "function（拆分字段）", "把 payload 拆成 rms / temp / state 四路，"
        "分别送仪表盘、曲线与数据库"],
       ["5", "ui_gauge × 2", "振动 RMS（0 ~ 0.1 g）、表面温度（0 ~ 60 ℃），三色分段"],
       ["6", "ui_text × 3", "设备上报状态、报警等级、维修建议"],
       ["7", "function（报警状态机）",
        "阈值判定 + **迟滞**：升级连续 3 帧确认、降级连续 5 帧确认"],
       ["8", "ui_chart × 2", "实时趋势（内存，保留 60 min）；"
        "历史曲线（SQLite 查询，每 30 s 刷新）"],
       ["9", "sqlite（prepared）",
        "`sqlquery` 必须选 **prepared**；SQL 固定，参数走 `msg.params`"],
       ["10", "file", "原始消息按行追加，作为离线留底"]],
      [Cm(1.0), Cm(4.2), Cm(11.2)], size=8.4, align_map=["c", "l", "l"])

h("6.2 数据库表结构", 2)
code_block([
    "CREATE TABLE readings (",
    "  id    INTEGER PRIMARY KEY AUTOINCREMENT,",
    "  ts    TEXT NOT NULL,   -- ISO8601 时间戳",
    "  node  TEXT,            -- 节点 ID",
    "  rms   REAL,            -- 振动加速度有效值 [g]",
    "  temp  REAL,            -- 表面温度 [℃]（DS18B20）",
    "  state TEXT             -- normal | warn | alarm",
    ");",
    "CREATE INDEX idx_readings_ts ON readings(ts);",
])
rich([("说明：", True),
      ("上表为**已实测**的结构。正式版在此基础上增加 `v_rms`（速度有效值 mm/s）、"
       "`tchip`（MPU6050 板载温度）、`fpeak`（基频带峰值频率）与 `score`（偏离度）四列；"
       "写入逻辑不变，只需在 INSERT 语句中补上对应参数。", False)])

h("6.3 仪表板", 2)
figure("fig4_dashboard.png", "图 S4　仪表板线框（Node-RED dashboard 实现）", 15.8)
bullets([
    "**页面一 · 实时状态**：振动 RMS 与表面温度仪表盘、设备上报状态、报警等级、维修建议、"
    "实时趋势曲线（保留 60 min）；",
    "**页面二 · 历史查询**：从 SQLite 读取最近 200 条绘制曲线，每 30 s 刷新，"
    "满足“历史可查 ≥ 4 周”的判据；",
    "**报警迟滞**：升级需连续 3 帧确认、降级需连续 5 帧确认，抑制阈值附近抖动；",
    "**建议与等级绑定**：正常 / 注意 / 报警三级各对应一条可执行建议；",
    "**访问方式**：编辑器 `http://localhost:1880`，仪表板 `http://localhost:1880/ui`；",
    "**流程文件** `flows.json` 与 `package.json` 提交 Git，队友 clone 后启动即复现。",
])

# ================================================================= 七
h("七、校准与测试记录", 1)
table("表 25　校准步骤与判据", ["项目", "方法", "判据"],
      [["加速度灵敏度与零偏", "重力翻转法：MPU6050 六个朝向各朝下，各采样 10 s，"
                       "最小二乘拟合", "拟合残差 ≤ 5 %"],
       ["加速度幅值（比对）", "与手持测振仪同点同时测量 3 个转速点", "偏差 ≤ 10 %"],
       ["表面温度", "冰水浴（0 ℃）与温水（约 45 ℃）与标准温度计比对", "偏差 ≤ 0.5 ℃"],
       ["芯片温度", "与 DS18B20 同步记录 30 min，给出固定偏置", "偏置稳定性 ± 1 ℃"],
       ["采样率", "连续采集 1 s，统计实际点数", "1000 ± 5"],
       ["频率准确度", "小音箱振膜 + 信号发生器输出 100 Hz 正弦", "谱峰误差 ≤ 1 %"],
       ["重复性", "同一工况连续测量 10 次", "变异系数 ≤ 3 %"]],
      [Cm(3.0), Cm(8.4), Cm(5.0)], size=8.5, align_map=["l", "l", "l"])
bullets([
    "每次测量当场填写台账「6-测试记录」，并保存原始数据文件；",
    "校准系数写入固件配置或 SD 卡配置文件，更换传感器后可重新标定；",
    "所有结论附不确定度说明，避免对测量能力作过度断言。",
])

# ================================================================= 八
h("八、演示台架与工况实验", 1)
h("8.1 台架搭法", 2)
bullets([
    "把 5 V 9225 风扇固定在木板或盒盖上，木板再固定到桌面；"
    "**风扇供电**：USB 公头插充电器，杜邦母头插风扇的 XH2.54-2P 公针，"
    "红线→5 V、黑线→GND，另外两根（D+/D−）悬空；风扇转速 **1700 RPM**（转频 28.3 Hz）；",
    "MPU6050 用双面胶或螺栓**贴紧**风扇外壳靠近轴承的位置（刚性安装）；",
    "DS18B20 探头贴在风扇外壳另一侧，与 MPU6050 保持几厘米间距，避免互相影响；",
    "不平衡：在叶轮同一位置贴配重片（2 g 与 5 g 两档），记录质量与位置；",
    "结构松动：松开固定风扇的一颗螺丝 1/4 圈或 1/2 圈，记录圈数；",
    "负载变化：用硬纸板遮挡进风口 30 % 与 60 %，记录遮挡比例。",
])
h("8.2 工况实验矩阵", 2)
table("表 26　工况变化对照实验矩阵", ["工况", "施加方式", "预期特征变化", "重复次数"],
      [["正常基线", "原始状态，稳定运行", "各指标稳定在 ±2σ 内", "5"],
       ["转子不平衡", "叶轮贴 2 g / 5 g 配重片", "基频带幅值与 RMS 上升", "5 / 档"],
       ["结构松动", "松开底座螺丝 1/4 圈、1/2 圈", "谐波成分上升，指标波动增大", "3 / 档"],
       ["负载变化", "遮挡进风口 30 % / 60 %", "RMS 与温度同时上升", "3 / 档"],
       ["转速变化", "USB 调压器降至 4 V 或换挡", "基频带峰值随转速移动（验证频带法的鲁棒性）",
        "3"]],
      [Cm(2.2), Cm(4.8), Cm(5.6), Cm(3.8)], size=8.5,
      align_map=["l", "l", "l", "c"])
para("每条记录须包含：时间戳、施加参数、全部特征值、原始 1024 点数据文件名。"
     "按工况分组统计均值与标准差，用箱线图与 ROC 曲线评估各判据的检出能力，"
     "并报告假阳性率。", indent=True)

# ================================================================= 九
h("九、需求与验收标准", 1)
para("本章给出本方案自身的功能性需求、非功能性需求与可测量的验收标准。"
     "第 3 列为优先级：M 必须 / S 应该 / C 可以。", indent=True)
table("表 27　功能性需求", ["编号", "需求描述", "优先级", "验证方法"],
      [["FR-01", "采集振动（I²C，1 kSPS，帧长 1024 点）、表面温度（1-Wire）"
                "与板载温度三路数据", "M", "数据记录与仪表板检查"],
       ["FR-02", "计算时域特征：RMS、峰值、速度有效值（10 ~ 250 Hz 积分）", "M",
        "注入已知信号比对"],
       ["FR-03", "完成 1024 点 FFT，给出主频与基频带（15 ~ 35 Hz）内峰值", "M",
        "音箱振膜 + 100 Hz 正弦验证"],
       ["FR-04", "支持开机基线自学习，并给出相对基线的偏离度", "M", "与离线计算比对"],
       ["FR-05", "输出“正常 / 注意 / 报警”三级状态与维护建议", "M",
        "人工制造工况变化验证"],
       ["FR-06", "通过 MQTT 上报特征值，周期 ≤ 10 s，事件即时上报", "M",
        "抓包与端到端计时"],
       ["FR-07", "断网时本地缓存，网络恢复后补传，不丢数据", "S", "拔网线测试"],
       ["FR-08", "提供中英双语仪表板，含实时值、历史趋势与频谱", "M",
        "界面走查 + 用户测试"],
       ["FR-09", "提供本地指示（板载 LED + 蜂鸣器），无网络时仍可判读", "S",
        "断网走查"],
       ["FR-10", "记录报警事件（含触发指标与时间戳）并可导出", "S", "导出并核对"],
       ["FR-11", "支持板载按键重新学习基线与静音", "C", "现场操作验证"]],
      [Cm(1.5), Cm(9.6), Cm(1.5), Cm(3.4)], size=8.5,
      align_map=["c", "l", "c", "l"])
table("表 28　非功能性需求", ["编号", "类别", "需求描述"],
      [["NFR-01", "安全", "全部电路 ≤ 5 V 直流（USB 供电），不接触市电；转动部件加防护"],
       ["NFR-02", "隐私", "不使用摄像头或麦克风，不采集任何可识别个人的数据"],
       ["NFR-03", "可靠性", "连续运行 ≥ 72 h 无重启；看门狗、掉电恢复、断网自愈"],
       ["NFR-04", "可维护性", "模块可现场插拔更换且免重新编译；基线可一键重学"],
       ["NFR-05", "成本", "单节点 ≤ ¥150（含耗材与备件）"],
       ["NFR-06", "可访问性", "仪表板中英双语；状态用文字 + 图标 + 颜色三重编码；字号可调"],
       ["NFR-07", "互操作性", "数据模型带版本号并有文档，可被第三方平台订阅复用"],
       ["NFR-08", "装配", "全程无需焊接：杜邦线直插 + 面包板电源轨分配"],
       ["NFR-09", "环境", "外壳防护 ≥ IP54，工作温度 0 ~ 50 ℃"]],
      [Cm(1.7), Cm(1.9), Cm(12.4)], size=8.5, align_map=["c", "c", "l"])
table("表 29　验收标准与验证方法", ["编号", "验收判据", "目标值", "验证方法"],
      [["AC-01", "振动采样率准确度", "误差 ≤ 0.5 %", "连续采集 1 s 统计实际点数（1000±5）"],
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
       ["AC-11", "单节点成本", "≤ ¥150", "按第二章清单核算（本版 ¥94 / ¥119）"],
       ["AC-12", "重复性", "变异系数 ≤ 3 %", "同一工况重复测量 10 次"],
       ["AC-13", "免焊接装配", "全部连接为插拔式，无需焊接", "装配过程走查与记录"]],
      [Cm(1.5), Cm(4.6), Cm(5.0), Cm(5.3)], size=8.4,
      align_map=["c", "l", "l", "l"])

# ================================================================= 十
h("十、与手册要求的逐条对照", 1)
table("表 30　手册要求与本方案的对应关系", ["手册要求", "本方案如何满足", "证据"],
      [["明确定义的客户需求、可衡量的验收标准", "见第九章表 27 ~ 表 29", "第九章"],
       ["可工作的集成原型（感应 → 处理 → 通信 → 输出）",
        "ESP32 单板完成全链路：采集 → FFT → 判别 → MQTT → 仪表板",
        "步骤 9 的端到端演示"],
       ["通常至少两个有意义的数据输入/传感器，并做校准",
        "三路：三轴加速度、表面温度（DS18B20）、板载温度（MPU6050），均可独立校准",
        "校准报告（重力翻转 / 水浴 / 比对）"],
       ["有意义的本地处理、过滤、融合、检测或决策逻辑",
        "本地完成 1024 点 FFT、特征提取、基线自学习与三级判别；"
        "两路温度与振动联合判据", "固件源码 + 单元测试记录"],
       ["适当的网络通信（Wi-Fi / MQTT）", "ESP32 内置 Wi-Fi + MQTT，10 s 周期 + 事件即时",
        "抓包记录 + 端到端延迟测试"],
       ["仪表板：实时与历史视图、报警与建议机制",
        "Node-RED dashboard：状态卡 + 24 h 趋势 + 频谱 + 分级报警与建议；SQLite 存历史",
        "界面走查 + 非项目成员可用性测试"],
       ["单元 / 集成 / 系统测试与实验证据", "第七章与第八章的校准与实验流程",
        "台账「6-测试记录」「7-工况实验记录」"],
       ["专业实践：版本控制、会议记录、问题跟踪、风险登记册、受控文档",
        "Git 仓库 + 受控文档台账（需求与验收、物料清单、风险登记册、责任矩阵、"
        "十二周计划）", "Git 仓库 + Excel 台账"],
       ["负责任设计：隐私、安全、伦理、可访问性",
        "无摄像头、无个人数据；全程 5 V USB 供电；仪表板中英双语、状态三重编码",
        "伦理 / 隐私 / 安全筛查表"],
       ["跨国部署：讨论在 UL 与 LICHPU 复制的可能性",
        "器件在两地均可采购；外壳防护、温度阈值与界面语言按环境调整", "见第十二章"]],
      [Cm(5.4), Cm(6.8), Cm(4.2)], size=8.1, align_map=["l", "l", "l"])

# ================================================================= 十一
h("十一、取舍与局限", 1)
para("手册的“优秀”评价档明确要求“诚实地陈述局限性”。"
     "以下内容应当直接写入最终报告的“讨论、局限性与经验教训”一节。", indent=True)
table("表 31　局限与应对", ["局限", "影响", "缓解措施与报告中的表述"],
      [["MPU6050 的 DLPF 带宽上限 260 Hz，采样率只能到 1 kSPS",
        "无法观测高频轴承共振，不能做包络解调分析",
        "本项目关注的 1×（约 24 Hz）、谐波、叶频（约 121 Hz）与 6205 轴承特征频率"
        "（BPFO 86 Hz / BPFI 131 Hz）全部低于 130 Hz，功能不受影响；"
        "报告中明确写出该带宽边界与它对可诊断故障类型的限制"],
       ["无转速传感器，不能做阶次归一化",
        "转速波动较大时判据会退化，仅在额定转速附近可靠",
        "改用“基频带 15 ~ 35 Hz 内最大峰值”作为指标；"
        "报告中写为“假定恒转速工况下的变化识别”"],
       ["无包络解调，对轴承早期点蚀不敏感", "无法在缺陷极早期发现轴承问题",
        "定位为识别不平衡、积尘、松动、负载变化四类——恰好是校园风机最常见故障"],
       ["本地指示降级为单色板载 LED", "现场只能靠闪烁频率区分三级状态",
        "配合蜂鸣器与仪表板；演示时投大屏更直观"],
       ["芯片温度（MPU6050 板载）受自热影响",
        "不能直接当作外壳表面温度使用",
        "仅作辅助通道与传感器健康指示；主温度仍用 DS18B20，并标定两者偏置"],
       ["Node-RED 依赖笔记本在线", "笔记本不在线时看不到实时界面",
        "节点侧 microSD 独立留存历史，补传后数据不丢"]],
      [Cm(3.6), Cm(4.2), Cm(8.6)], size=8.2, align_map=["l", "l", "l"])

# ================================================================= 十二
h("十二、跨国复制与部署要点", 1)
table("表 32　UL 与 LICHPU 复制差异与对策", ["维度", "差异", "对策"],
      [["采购与成本", "中国采购约 ¥94；爱尔兰本地分销约 €45 ~ €60",
        "集中采购后随身携带；模块化器件在两地均易采购，复制门槛低"],
       ["供电", "爱尔兰 230 V / 中国 220 V，均为 50 Hz；本方案使用 USB 5 V",
        "只需符合当地认证的 USB 充电器（CE / CCC），节点侧接口不变"],
       ["气候与环境", "爱尔兰湿冷多雨易凝露；中国夏季高温高湿",
        "IP54 外壳 + 透气防水阀；接线端子防潮；温度阈值用与环境温差而非绝对温度"],
       ["网络", "两地校园 Wi-Fi 均可用 2.4 GHz 信道 1 ~ 13",
        "节点不硬编码信道；MQTT Broker 在两地各部署或使用云端"],
       ["设备与基线", "两地风机品牌、叶片数与转速不同",
        "基线自学习使阈值自适应；基频带法不依赖精确转速，复制时无需重新标定"],
       ["语言与可访问性", "两地语言不同",
        "仪表板中英双语；状态用文字 + 图标 + 颜色三重编码"],
       ["法规与数据保护", "爱尔兰适用 GDPR；中国适用 PIPL",
        "不采集个人数据，仅聚合设备物理量，满足数据最小化；数据限域存储、保留期 ≤ 4 周"]],
      [Cm(2.4), Cm(5.8), Cm(8.2)], size=8.4, align_map=["l", "l", "l"])

# ================================================================= 十三
h("十三、后续扩展方向", 1)
para("若中期评审后时间仍有富余，可按下列顺序逐步增强；每一步都是独立可交付的增量，"
     "且都不需要焊接。", indent=True)
table("表 33　可选增强项（时间允许时按顺序实施）", ["顺序", "增强项", "增加成本", "增加工作量"],
      [["1", "加装光电/霍尔转速传感器，用中断 + micros() 测周期", "¥6", "1 ~ 2 天"],
       ["2", "加入转速分档基线与阶次谱", "¥0", "2 ~ 3 天"],
       ["3", "把 MPU6050 换成 ICM-42688-P（ODR 达 32 kHz），恢复高频带宽并做包络解调",
        "¥30", "2 ~ 3 天"],
       ["4", "利用 MPU6050 自带的陀螺仪，增加“安装松动/姿态变化”检测通道", "¥0",
        "1 ~ 2 天"],
       ["5", "服务端升级为 InfluxDB + Grafana", "¥0", "1 ~ 2 天"],
       ["6", "部署到校园真实风机做 4 周试运行", "¥0（复用现有节点）", "持续运行 + 每周巡检"]],
      [Cm(1.3), Cm(8.7), Cm(2.9), Cm(3.0)], size=8.5,
      align_map=["c", "l", "c", "c"])

# ================================================================= 附录
h("附录 A　开发环境搭建与故障速查（本方案实测）", 1)
para("本附录记录在国内网络环境下**实际搭建并跑通**的软件版本与遇到的故障，"
     "供队友复现与答辩答疑使用。所有条目均为实测结论，不是推测。", indent=True)

h("A.1 软件清单", 2)
table("表 34　实测通过的软件与版本", ["软件 / 包", "版本", "用途"],
      [["Arduino IDE", "2.3.x", "编译与烧录"],
       ["esp32 板支持包", "**2.0.17**", "见 A.2 第 2 条：不要用 3.x"],
       ["CH341SER 驱动", "WCH 官方", "CH340 串口（设备管理器显示 USB-SERIAL CH340）"],
       ["PubSubClient", "2.8.0", "ESP32 侧 MQTT 客户端"],
       ["ArduinoJson", "7.4.3", "JSON 序列化（也可手写 JSON 字串，不依赖版本）"],
       ["Node.js", "24.21.0 LTS", "运行 Node-RED"],
       ["Node-RED", "5.0.8", "服务端编排"],
       ["node-red-dashboard", "3.6.6", "仪表板"],
       ["node-red-contrib-aedes", "1.4.0", "MQTT broker"],
       ["node-red-node-sqlite", "2.0.1 + sqlite3 6.0.1", "历史存储"]],
      [Cm(4.0), Cm(3.6), Cm(8.8)], size=8.4, align_map=["l", "c", "l"])

h("A.2 七个高频故障与对策", 2)
table("表 35　故障速查（按实际遇到顺序）", ["现象", "根因", "对策"],
      [["开发板包安装失败，报 `Download failed … 185.199.109.133`",
        "`release-assets.githubusercontent.com` 在国内不可达",
        "把 `https://github.com/…` 换成 `https://dl.espressif.com/github_assets/…`，"
        "预下载到 `%LOCALAPPDATA%\\Arduino15\\staging\\packages\\`，再点安装即走本地缓存"],
       ["选最新版要下载约 1.9 GB",
        "3.x 含 C3/C5/C6/H2/P4/S2/S3 的 RISC-V 工具链与预编译库",
        "装 **2.0.17**（约 828 MB）；且本方案代码按 2.x API 编写，版本与文档一致"],
       ["编译报 `collect2.exe: error: ld returned 5`；或板子无限重启 "
        "`rst:0x3 (SW_RESET)`",
        "烧录或编译进程被中途终止，固件/目标文件不完整",
        "完整重跑；**不要对 arduino-cli 的输出管道做提前截断**，那会杀掉子进程"],
       ["`npm` 报“无法加载文件 … 在此系统上禁止运行脚本”",
        "PowerShell 执行策略为 Restricted",
        "`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`"],
       ["`sqlite3` 安装失败，报 node-gyp / 找不到 Python",
        "预编译包从 GitHub 下载超时，回退本地编译又缺 Python 与 MSVC",
        "用镜像下载 `sqlite3-vX-napi-v6-win32-x64.tar.gz`（N-API 跨 Node 版本兼容），"
        "解压到 `node_modules/sqlite3/build/Release/`"],
       ["Node-RED 里 MQTT 订阅**收不到数据**，且无任何报错",
        "`mqtt in` 的 broker 指向了 **aedes broker 节点**（服务器），"
        "而非 **mqtt-broker 配置节点**（客户端）",
        "新建 `mqtt-broker` 配置节点指向 `127.0.0.1:1883`，让 `mqtt in` 引用它"],
       ["sqlite 节点不写数据、也不报错",
        "`sqlquery` 只接受 `msg.topic` / `fixed` / `prepared`，填错则**静默不执行**",
        "固定 SQL + `msg.params` 时必须选 **prepared**"]],
      [Cm(4.6), Cm(4.4), Cm(7.4)], size=8.0, align_map=["l", "l", "l"])

h("A.3 网络与端口", 2)
bullets([
    "ESP32 与运行 Node-RED 的笔记本必须**同一网段**（实测：热点 10.77.48.0/24、"
    "笔记本 10.77.48.14、ESP32 由 DHCP 分配 10.77.48.96）；",
    "校园网 `HPU_WiFi` 采用网页 / 802.1X 认证，ESP32 无法接入——"
    "**必须用手机热点且工作在 2.4 GHz**；",
    "端口：Node-RED 编辑器 **1880**、MQTT broker **1883**；",
    "Windows 防火墙通常已有 Node.js 的入站规则；若 ESP32 连不上 broker，先查 1883 是否被拦。",
])

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(12)
setfont(p.add_run("（实施方案完）"), size=10, ea="黑体",
        color=RGBColor(0x70, 0x70, 0x70))
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.save(OUT)
print("saved:", OUT)
