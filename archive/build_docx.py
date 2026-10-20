# -*- coding: utf-8 -*-
"""生成《震动状态监测器（选题 8.1）设计方案》Word 文档。"""
import datetime
import os

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
OUT = os.path.join(BASE, "震动状态监测器设计方案.docx")

DARK = RGBColor(0x1F, 0x3B, 0x63)
BLACK = RGBColor(0x20, 0x20, 0x20)
ACCENT = RGBColor(0xA0, 0x3A, 0x1E)

doc = Document()

# ----------------------------------------------------------------- 页面设置
sec = doc.sections[0]
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.8)
sec.right_margin = Cm(2.8)
CONTENT_W = Cm(21.0 - 5.6)


# ----------------------------------------------------------------- 基础工具
def setfont(run, size=10.5, bold=False, ascii_font="Times New Roman",
            ea="宋体", color=BLACK, italic=False):
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
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
doc.styles["Normal"].paragraph_format.line_spacing = 1.35
doc.styles["Normal"].paragraph_format.space_after = Pt(4)

for lvl, (sz, ea) in {1: (15, "黑体"), 2: (12.5, "黑体"), 3: (11, "黑体")}.items():
    st = fix_style(f"Heading {lvl}", sz, ea=ea, ascii_font="Arial", bold=True,
                   color=DARK)
    st.paragraph_format.space_before = Pt(10 if lvl == 1 else 8)
    st.paragraph_format.space_after = Pt(5)
    st.paragraph_format.line_spacing = 1.2


def para(text="", *, size=10.5, bold=False, align="left", ea="宋体",
         color=BLACK, first_indent=True, space_after=4, ascii_font="Times New Roman"):
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    if first_indent:
        pf.first_line_indent = Pt(size * 2)
    r = p.add_run(text)
    setfont(r, size=size, bold=bold, ea=ea, color=color, ascii_font=ascii_font)
    return p


def rich(parts, *, size=10.5, align="left", first_indent=True, space_after=4):
    """parts: [(text, bold), ...]"""
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    p.paragraph_format.space_after = Pt(space_after)
    if first_indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    for t, b in parts:
        setfont(p.add_run(t), size=size, bold=b)
    return p


def h(text, level=1):
    p = doc.add_heading("", level=level)
    ea = "黑体"
    setfont(p.add_run(text), size={1: 15, 2: 12.5, 3: 11}[level], bold=True,
            ea=ea, ascii_font="Arial", color=DARK)
    return p


def bullets(items, *, size=10.5, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.3
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


def set_cell(cell, text, *, size=9, bold=False, align="left", ea="宋体"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER}[align]
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.15
    setfont(p.add_run(text), size=size, bold=bold, ea=ea)


def table(caption, headers, rows, widths, *, size=9, align_map=None,
          header_fill="DEEBF7"):
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(6)
    cp.paragraph_format.space_after = Pt(2)
    cp.paragraph_format.keep_with_next = True
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
            a = "left"
            if align_map and align_map[i] == "c":
                a = "center"
            set_cell(cells[i], str(v), size=size, align=a)
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = widths[i]
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def figure(fname, caption, width_cm):
    doc.add_picture(os.path.join(FIG, fname), width=Cm(width_cm))
    pic_p = doc.paragraphs[-1]
    pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p.paragraph_format.space_before = Pt(6)
    pic_p.paragraph_format.space_after = Pt(2)
    pic_p.paragraph_format.keep_with_next = True
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(8)
    setfont(cp.add_run(caption), size=9, ea="黑体", color=RGBColor(0x50, 0x50, 0x50))


# ----------------------------------------------------------------- 页脚页码
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
setfont(footer_p.add_run("震动状态监测器设计方案（选题 8.1）　第 "), size=8, ea="宋体")
r = footer_p.add_run()
f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
r._r.append(f1); r._r.append(it); r._r.append(f2)
setfont(r, size=8)
setfont(footer_p.add_run(" 页"), size=8, ea="宋体")

TODAY = datetime.date.today()

# ================================================================= 封面块
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("课程设计小组方案"), size=12, bold=True, ea="黑体", color=ACCENT)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("震动状态监测器"), size=26, bold=True, ea="黑体", color=DARK)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(10)
setfont(p.add_run("设 计 方 案"), size=14, bold=True, ea="黑体", color=DARK)

table("", ["项目", "内容"],
      [["选题编号", "8.1"],
       ["选题名称", "震动状态监测器（旋转机械振动状态监测与故障预警装置）"],
       ["方案版本", "V1.0（方案论证稿）"],
       ["小组名称 / 成员", "　　　　　　　　　　　　　　　　　　（按实际填写）"],
       ["指导教师", "　　　　　　　　　　"],
       ["日期", TODAY.strftime("%Y 年 %m 月 %d 日")]],
      [Cm(3.6), Cm(11.8)], size=10, align_map=["c", "l"])

h("摘要", 1)
para("本方案针对课程设计选题 8.1“震动状态监测器”提出一套完整、可落地的实现方案。监测器以 "
     "STM32F407VET6 为主控，采用数字 MEMS 加速度计 ADXL345（SPI，±16 g，ODR 3.2 kHz）作为主测量通道，"
     "并预留 IEPE 压电加速度计调理通道作为扩展，实现振动信号的采集、预处理、2048 点 FFT 频谱分析、"
     "时域与频域特征提取（RMS、峰值、峰值因子、峭度、速度有效值、包络谱），"
     "按“正常 / 注意 / 报警”三级判据输出状态结论，通过 OLED 本地显示、蜂鸣器与三色 LED 声光报警、"
     "microSD 数据记录与 ESP8266 WiFi 数据上报，配套上位机软件实现趋势查看。"
     "方案给出了完整的硬件框图、软件流程、关键器件选型对比、判别阈值、标定与误差分析方法、"
     "测试验证方案、进度分工与成本预算，单台样机成本约 215 元（含 IEPE 扩展约 305 元）。", align="justify")

para("说明：本方案依据选题名称编制。若课程任务书对量程、频带、通信方式、供电方式或成本上限另有明确规定，"
     "请以任务书为准，并按附录 B 给出的对应调整建议替换相关章节。", size=9.5, color=RGBColor(0x70, 0x70, 0x70))

# ================================================================= 一
h("一、选题背景与设计目标", 1)
h("1.1 选题背景与意义", 2)
para("电动机、风机、水泵、齿轮箱、机床主轴等旋转机械在长期运行中，会因轴承磨损与点蚀、转子不平衡、"
     "联轴器不对中、齿轮断齿、地脚松动、部件共振等原因产生异常振动。这些缺陷在形成早期往往只表现为"
     "振动信号中微弱的周期性冲击或特定频率成分的能量上升，肉眼、听感乃至温度、电流等特征量都难以察觉，"
     "而振动信号频带宽、响应快、对早期缺陷最敏感，因此振动监测是设备状态监测与故障诊断中最有效、"
     "最成熟的技术手段。", align="justify")
para("本选题面向中小型旋转机械，研制一台便携式/固定式震动状态监测器：以加速度传感器拾取振动信号，"
     "在嵌入式平台上完成采集、频谱分析与特征提取，依据国标/ISO 振动烈度分级思想输出设备健康状态，"
     "并在越限时声光报警、记录故障数据、上传上位机。装置可直接用于实验室转子试验台与课程演示，"
     "也具备一定的工程实用价值。", align="justify")

h("1.2 设计目标与技术指标", 2)
para("监测器的总体目标是：测得准、算得快、判得稳、看得见、报得出、存得住。具体技术指标如表 1。",
     align="justify")
table("表 1　震动状态监测器技术指标", ["项目", "指标要求", "备注 / 依据"],
      [["测量物理量", "振动加速度（并换算速度有效值、位移）", "加速度积分得速度"],
       ["量程", "±16 g（可切换 ±2 / ±4 / ±8 g）", "ADXL345 四档可选"],
       ["幅值分辨率", "3.9 mg/LSB（±16 g 档，全分辨率 13 bit）", "对应约 0.038 m/s²"],
       ["频率范围", "5 Hz ~ 1.6 kHz（−3 dB）", "兼顾轴承特征频率"],
       ["采样率", "3.2 kSPS（可切 1.6 / 0.8 kSPS）", "与 ODR 联动"],
       ["帧长 / 分辨率", "N = 2048 点，Δf = 1.5625 Hz", "可切 4096 点，Δf = 0.78 Hz"],
       ["特征参数", "RMS、峰值、峰峰值、峰值因子、峭度、速度有效值、包络谱峰值",
        "时域 + 频域 + 包络域"],
       ["状态分级", "正常 / 注意 / 报警 三级", "参考 ISO 10816-3 振动烈度分区"],
       ["报警响应时间", "≤ 1 s（含 3 帧确认）", "1 s 内完成声光报警"],
       ["本地显示", "0.96 英寸 OLED 128×64：数值页 / 波形页 / 频谱页 / 设置页", "I²C 接口"],
       ["数据存储", "microSD（≥ 8 GB），CSV + 二进制原始数据", "FATFS 文件系统"],
       ["通信方式", "ESP8266 WiFi（TCP / MQTT），预留 UART / RS485", "特征值 JSON 上报"],
       ["供电与续航", "3.7 V / 2000 mAh 锂电池，续航 ≥ 8 h", "支持 USB 充电"],
       ["整机成本", "≤ 300 元 / 台（不含仪器仪表）", "见表 4"],
       ["工作温度", "0 ~ 50 ℃", "实验室与一般工业环境"]],
      [Cm(3.2), Cm(7.6), Cm(4.6)], size=9, align_map=["c", "l", "l"])

# ================================================================= 二
h("二、需求分析与总体方案", 1)
h("2.1 功能需求分析", 2)
bullets([
    [("信号采集与调理：", True), ("拾取三轴振动加速度，完成交流耦合、放大、抗混叠滤波与模数转换，"
                            "保证 5 Hz ~ 1.6 kHz 内有平坦幅频特性和足够动态范围。", False)],
    [("数据处理与特征提取：", True), ("本地完成去趋势、加窗、FFT，计算时域、频域与包络域特征值。", False)],
    [("状态判别与报警：", True), ("依据阈值与基线给出三级状态结论，越限时声光报警并锁存事件。", False)],
    [("人机交互：", True), ("OLED 实时显示特征值与波形/频谱，按键完成量程、阈值、页面的设置。", False)],
    [("数据存储：", True), ("按时间命名记录特征值与故障快照，掉电不丢失，可用读卡器导出。", False)],
    [("数据通信：", True), ("WiFi 上报特征值至上位机，形成振动趋势曲线与报警记录。", False)],
    [("电源管理：", True), ("锂电池供电与充电管理、低功耗策略、电量指示。", False)],
    [("标定与自检：", True), ("支持灵敏度标定系数写入，上电自检传感器在线状态。", False)],
])

h("2.2 系统总体方案", 2)
para("系统采用“传感器 + 信号调理 + 嵌入式主控 + 人机交互 + 通信 + 电源”的模块化结构。"
     "主测量通道选用数字 MEMS 加速度计，免去模拟调理环节，接口简单、成本低、可靠性高；"
     "同时预留一路 IEPE 压电加速度计调理通道，用于后续提升带宽与动态范围。"
     "主控在本地完成全部信号处理，仅把特征值上传，既降低通信带宽占用，也保证断网时监测不中断。"
     "系统总体框图见图 1。", align="justify")
figure("fig1_system.png", "图 1　震动状态监测器系统总体框图", 15.2)

h("2.3 关键环节方案对比与选型", 2)
para("针对传感器、主控、显示与通信四个关键环节，对候选方案进行对比论证，结果见表 2。", align="justify")
table("表 2　关键环节方案对比与选型", ["环节", "候选方案", "主要优点", "主要不足", "结论"],
      [["振动\n传感器",
        "A. 数字 MEMS 加速度计 ADXL345",
        "SPI/I²C 数字输出免调理；±16 g；ODR 达 3.2 kHz；成本低、资料丰富",
        "带宽上限 1.6 kHz；噪声密度约 290 µg/√Hz，不及压电式",
        "★ 采用（主通道）"],
       ["", "B. 模拟 MEMS 加速度计 + 外部高速 ADC",
        "带宽可达 10 kHz 以上，动态范围较好",
        "需外部 ADC、基准与模拟前端，电路复杂、成本翻倍",
        "备选"],
       ["", "C. IEPE 压电加速度计 + 恒流源 + 高速 ADC",
        "工业标准；带宽宽、噪声低、量程大，可测轴承高频共振",
        "需 24 V 恒流源与电荷/电压调理；单价高，标定复杂",
        "★ 采用（扩展通道）"],
       ["主控",
        "A. STM32F407VET6",
        "168 MHz，硬件 FPU + DSP 指令，192 KB SRAM，3×12-bit ADC，SDIO/USB/以太网接口齐全",
        "单价高于 F103，封装需自制或核心板",
        "★ 采用"],
       ["", "B. STM32F103C8T6",
        "价格极低、生态成熟",
        "无 FPU，20 KB SRAM，2048 点浮点 FFT 与多特征计算吃力",
        "不采用"],
       ["", "C. ESP32",
        "集成 WiFi，算力较强，双核",
        "片内 ADC 噪声大、非线性明显；WiFi 实时性影响采样节拍",
        "不采用（仅作通信模块）"],
       ["显示",
        "A. OLED 0.96″ SSD1306（I²C）",
        "功耗低（约 15 mA）、对比度高、占用引脚少、驱动成熟",
        "尺寸小，单屏信息量有限，需分页显示",
        "★ 采用"],
       ["", "B. TFT LCD 2.4″（SPI）",
        "色彩与信息量大，可画波形",
        "功耗高（80 mA 以上），占用引脚多，成本较高",
        "备选"],
       ["通信",
        "A. ESP8266-01S（UART AT）",
        "成本低、协议栈完整，可 TCP/MQTT，覆盖实验室范围",
        "功耗峰值较高，需注意电源去耦",
        "★ 采用"],
       ["", "B. HC-05 蓝牙串口", "配对简单、功耗低",
        "传输距离短（约 10 m），仅点对点，不适合多机汇聚", "备选"],
       ["", "C. RS485（MAX485）", "抗干扰强、传输距离可达 1200 m，工业通用",
        "需额外布线，不适合便携演示", "预留接口"]],
      [Cm(1.6), Cm(3.6), Cm(4.0), Cm(4.0), Cm(2.2)], size=8.5,
      align_map=["c", "l", "l", "l", "c"])

# ================================================================= 三
h("三、硬件设计", 1)
h("3.1 传感器与信号调理电路", 2)
h("3.1.1 主通道：数字 MEMS 加速度计", 3)
bullets([
    "器件：ADXL345（三轴数字加速度计），量程软件可选 ±2/±4/±8/±16 g，全分辨率模式 13 bit，"
    "±16 g 档灵敏度 3.9 mg/LSB。",
    "接口：SPI 4 线（SCLK / MISO / MOSI / CS），时钟不超过 5 MHz，接线短、抗干扰好；"
    "采用 FIFO 流模式（32 级）配合 INT1 数据就绪中断，避免高频轮询。",
    "配置：ODR = 3200 Hz（带宽 ≈ ODR/2 = 1600 Hz），启用全分辨率与低噪声模式；"
    "切换采样率时同步调整 ODR，保证数字通道不产生混叠。",
    "安装：优先采用螺栓或粘接刚性安装，传感器安装面需平整；监测高频成分时避免使用磁座"
    "（磁座会引入约 1 kHz 附近的安装共振）。",
])
h("3.1.2 扩展通道：IEPE 压电加速度计调理", 3)
bullets([
    "恒流源：以运放 + P-MOS 或 LM334 构成 4 mA / 24 V 恒流源，为 IEPE 传感器供电并拾取交流分量。",
    "隔直与放大：隔直电容去除约 12 V 直流偏置；两级同相放大，增益 1×/10×/100× 由模拟开关或"
    "跳线选择，用于适配不同灵敏度传感器（如 100 mV/g 与 10 mV/g）。",
    "抗混叠：二阶 Sallen-Key 巴特沃斯低通，fc ≈ 1.5 kHz，与 3.2 kSPS 采样率配合，"
    "在奈奎斯特频率 1.6 kHz 处提供约 −12 dB 衰减。",
    "偏置与保护：输出偏置至 1.65 V（VREF/2）以匹配 ADC 输入范围，输入端加 TVS 与限流电阻保护。",
])
h("3.1.3 抗混叠策略", 3)
para("数字 MEMS 通道依靠器件内部按 ODR 设定的低通特性限制带宽，采样率与 ODR 严格联动；"
     "IEPE 模拟通道设置二阶有源低通；软件层再对 1.5 kHz 以上残余分量做一次数字平均或"
     "FIR 半带滤波，形成“器件—模拟—数字”三级抗混叠，避免高频振动折叠到低频特征频段造成误判。",
     align="justify")

h("3.2 主控与数据采集", 2)
bullets([
    "最小系统：STM32F407VET6 + 8 MHz 晶振（PLL 至 168 MHz）+ SWD 调试口 + 复位与 BOOT 电路。",
    "采集链路：主通道由 SPI1 + DMA 从 ADXL345 FIFO 突发读取三轴数据；扩展通道由 TIM2 定时触发"
    "ADC1，DMA 双缓冲搬运 2048 点，采样时刻由定时器硬件决定，抖动小于 1 个采样周期。",
    "运算：使用 CMSIS-DSP 库的 arm_rfft_fast_f32 完成 2048 点实数 FFT，配合 arm_max_f32、"
    "arm_mean_f32 等完成特征计算，单帧处理时间远小于 1 s 的帧周期。",
    "采样率核算：fs = 3200 Hz，1 s 内可完成 1.56 帧（50 % 重叠时约 3 帧），"
    "既保证报警响应 ≤ 1 s，又给频谱平均留出余量。",
])
table("表 3　主控资源预算（STM32F407VET6，SRAM 192 KB / Flash 1 MB）",
      ["资源项", "占用", "说明"],
      [["ADC/SPI 双缓冲原始数据", "2 × 2048 × 4 B = 16 KB", "双缓冲保证采集与处理并行"],
       ["FFT 输入/输出缓冲", "2048 × 4 B × 2 = 16 KB", "单精度浮点，就地运算可再省 8 KB"],
       ["窗函数与频谱幅度缓冲", "约 12 KB", "Hanning 表 + 幅值谱"],
       ["特征值与统计缓冲", "约 8 KB", "含包络谱与滑动帧缓存"],
       ["协议/文件系统缓冲", "约 12 KB", "FATFS + WiFi 收发缓冲"],
       ["合计", "约 64 KB（占 33 %）", "余量充足，可支持 4096 点长帧升级"]],
      [Cm(4.2), Cm(5.2), Cm(6.0)], size=9, align_map=["l", "c", "l"])

h("3.3 存储与人机交互", 2)
bullets([
    "存储：microSD 卡（SDIO 4 位或 SPI 模式）+ FATFS，文件按监测器编号与日期命名"
    "（如 VIB01_20250315.csv），每秒写入一行特征值，越限时另存一帧原始数据二进制文件作为故障快照。",
    "显示：0.96″ OLED（SSD1306，I²C）四页界面，依次为“实时数值与状态”“时域波形”"
    "“频谱与主频”“设置与标定”。",
    "按键：3 个轻触按键（上翻 / 下翻 / 确认），长按进入设置，完成量程、采样率、"
    "阈值与通信参数配置。",
    "声光报警：有源蜂鸣器（S8050 三极管驱动，续流二极管保护）+ 红/黄/绿三色 LED，"
    "分别对应报警 / 注意 / 正常，串限流电阻。",
])

h("3.4 通信与上位机", 2)
bullets([
    "硬件：ESP8266-01S 通过 UART2（115200 bps）与主控连接。",
    "协议：采用行式 JSON 帧上报，字段包括设备号、时间戳、RMS、峰值、峭度、速度有效值、"
    "主频与状态等级；同时提供查询与阈值下发指令。",
    "上位机：可用 Python（PyQt5 + pyqtgraph）或 MQTT + Node-RED 实现，功能包括实时数值与"
    "趋势曲线、报警记录列表、历史数据回放与报表导出。",
    "断网策略：本地 SD 卡先写后传，WiFi 断连时进入缓存队列并在恢复后补传，"
    "保证监测连续性。",
])

h("3.5 电源方案", 2)
bullets([
    "供电链：3.7 V / 2000 mAh 锂电池 → TP4056 充电管理（含 DW01 过充过放保护）→ "
    "低压差 LDO（ME6211C33，压差约 100 mV）输出 3.3 V。",
    ("说明：不宜使用压差约 1.1 V 的 AMS1117 —— 电池电压降至 3.5 V 以下时会跌出稳压区，"
     "导致系统复位，这是常见的设计陷阱。", False),
    "扩展通道电源：MT3608 升压模块输出 24 V，为 IEPE 恒流源供电，仅在扩展通道启用时开启以省电。",
    "功耗估算：主控约 60 mA + OLED 15 mA + 传感器 0.2 mA + WiFi 平均 70 mA ≈ 145 mA；"
    "按电池可用容量 1600 mAh 计，续航约 11 h，满足 ≥ 8 h 指标；关闭 WiFi 可延长至 20 h 以上。",
    "低功耗策略：空闲时降频、屏幕自动息屏、WiFi 定时间隔上报（如 10 s 一次）。",
])

h("3.6 元器件清单与成本预算", 2)
table("表 4　元器件清单与成本预算（单台，参考价）",
      ["序号", "名称", "型号规格", "数量", "单价/元", "小计/元"],
      [["1", "主控核心板", "STM32F407VET6 最小系统板", "1", "55", "55"],
       ["2", "加速度传感器", "ADXL345 模块（SPI/I²C）", "1", "12", "12"],
       ["3", "显示模块", "OLED 0.96″ SSD1306 I²C", "1", "12", "12"],
       ["4", "存储", "microSD 卡座 + 8 GB 卡", "1", "25", "25"],
       ["5", "无线通信", "ESP8266-01S 模块", "1", "12", "12"],
       ["6", "电池与充电", "2000 mAh 锂电 + TP4056 + 保护板", "1", "35", "35"],
       ["7", "电源器件", "ME6211C33 LDO、电容电感、接插件、线材", "1", "24", "24"],
       ["8", "人机交互件", "蜂鸣器、三色 LED、轻触按键、洞洞板/PCB", "1", "40", "40"],
       ["", "小计（基础方案）", "", "", "", "215"],
       ["9", "（选）压电加速度计", "IEPE 型，100 mV/g", "1", "60", "60"],
       ["10", "（选）调理电路器件", "恒流源、运放、精密电阻电容、24 V 升压", "1", "30", "30"],
       ["", "小计（含扩展通道）", "", "", "", "305"]],
      [Cm(1.2), Cm(2.8), Cm(5.6), Cm(1.2), Cm(1.9), Cm(1.9)], size=9,
      align_map=["c", "l", "l", "c", "c", "c"])
para("说明：以上为淘宝/立创商城常见零售价，PCB 打样、示波器与振动台等仪器不计入；"
     "若采用 PCB 自制替代核心板与模块，基础方案成本可再降 30 ~ 40 元。", size=9.5, align="justify")

figure("fig2_signal.png", "图 2　信号调理、采集与处理链路", 15.2)

# ================================================================= 四
h("四、软件设计", 1)
h("4.1 软件总体流程", 2)
para("软件采用前后台结构：主循环负责任务调度与运算，中断服务程序只做最短的数据搬运与标志置位，"
     "保证采样节拍严格、处理不丢帧。程序流程见图 3。", align="justify")
figure("fig3_software.png", "图 3　软件主流程", 13.0)

h("4.2 数据采集与预处理", 2)
bullets([
    "采样参数：fs = 3200 Hz，帧长 N = 2048（时窗 0.64 s），帧重叠 50 %，等效每秒输出 3 帧结果。",
    "预处理顺序：去均值（消零偏）→ 最小二乘去趋势（消温漂与基线漂移）→ Hanning 加窗"
    "抑制频谱泄漏，并乘以幅值恢复系数 2.0 使谱峰幅值还原。",
    "单位换算：a[m/s²] = 9.80665 × a[g]；速度谱由加速度谱积分 A(f)/(2πf) 得到，"
    "积分区间取 10 ~ 1000 Hz 以避免低频积分发散。",
])

h("4.3 特征提取算法", 2)
rich([("时域特征：", True),
      ("均方根值 X_rms = √( (1/N) Σ x_i² )，反映振动能量，是状态判别的首选指标；"
       "峰值 X_p = max|x_i|，峰峰值 X_pp = X_max − X_min；"
       "峰值因子 C_f = X_p / X_rms（正常约 3 ~ 4，出现冲击故障时升高至 5 以上）；"
       "峭度 K = (1/N)Σ(x_i − μ)⁴ / σ⁴（正常约 3，早期冲击故障时明显大于 4）。", False)])
rich([("频域特征：", True),
      ("2048 点 FFT 得到幅值谱后，提取主频 f_p、谱重心频率、各次谐波幅值、"
       "以及 0.5×、1×、2×、3× 转频分量比值，用于识别不平衡（1× 突出）、"
       "不对中（2× 突出）、松动（多阶谐波与边频）等典型故障。", False)])
rich([("速度有效值：", True),
      ("v_rms = √( Σ (A_k / 2π f_k)² / 2 )（mm/s），在 10 ~ 1000 Hz 内积分，"
       "该指标与 ISO 10816 振动烈度直接对应，便于与标准限值比较。", False)])
rich([("包络解调特征：", True),
      ("对 500 ~ 1500 Hz 频带（轴承共振区）带通滤波后做 Hilbert 变换取包络，再对包络做 FFT，"
       "在包络谱中寻找轴承故障特征频率及其谐波。以 6205 深沟球轴承（球数 N = 9，"
       "球径 d = 7.94 mm，节圆直径 D = 38.5 mm，接触角 α ≈ 0°，转速 1500 r/min，"
       "转频 f_r = 25 Hz）为例：外圈故障频率 BPFO = (N/2)·f_r·(1 − (d/D)cosα) ≈ 89.3 Hz，"
       "内圈故障频率 BPFI ≈ 135.7 Hz，滚动体故障频率 BSF ≈ 58.0 Hz，"
       "保持架故障频率 FTF ≈ 9.9 Hz。三者均远低于 1.6 kHz 分析上限，"
       "验证了 3.2 kSPS 采样率对本类轴承故障的覆盖能力。", False)])
rich([("分辨率校核：", True),
      ("Δf = fs/N = 1.5625 Hz；对 FTF ≈ 9.9 Hz 及低速轴成分，分辨率略显不足，"
       "此时自动切换到 N = 4096（Δf = 0.78 Hz）或采用包络谱细化，"
       "确保特征频率误差小于 1 个频率分辨率。", False)])

h("4.4 状态判别与报警策略", 2)
para("判别采用“多判据 + 分级 + 迟滞”的策略，避免单一指标误判与阈值附近的频繁抖动：", align="justify")
bullets([
    "基线自学习：开机空载（或设备正常工况）采集 30 s，统计 RMS、峭度与速度有效值的均值 μ₀ 与"
    "标准差 σ₀，自动生成注意阈值 μ₀ + 3σ₀ 与报警阈值 μ₀ + 6σ₀，并允许人工修改。",
    "多判据表决：速度有效值（对应振动烈度）、加速度 RMS、峭度三项中任意一项进入更高等级，"
    "即判定状态升级，兼顾能量型与冲击型故障。",
    "迟滞与消抖：连续 3 帧越限才确认报警；报警恢复需连续 5 帧回到正常区间，防止临界抖动。",
    "报警动作：蜂鸣器鸣响 + 红灯点亮 + OLED 显示状态与主要特征值 + 保存故障快照 + 立即 WiFi 上报。",
])
table("表 5　状态判别阈值（三级判据，示例配置）",
      ["判别量", "正常（绿）", "注意（黄）", "报警（红）"],
      [["速度有效值 v_rms（mm/s，10 ~ 1000 Hz）",
        "< 0.71", "0.71 ~ 1.8", "> 1.8（≥ 4.5 为危险）"],
       ["加速度 RMS（g）", "< 0.5（或 < μ₀+3σ₀）", "0.5 ~ 1.5", "> 1.5（或 > μ₀+6σ₀）"],
       ["峭度 K", "< 4.0", "4.0 ~ 5.0", "> 5.0"],
       ["峰值因子 C_f", "< 4.0", "4.0 ~ 5.0", "> 5.0"],
       ["包络谱特征频率幅值", "低于基线 3 倍", "基线 3 ~ 6 倍", "高于基线 6 倍"]],
      [Cm(5.4), Cm(3.4), Cm(3.3), Cm(3.3)], size=9, align_map=["l", "c", "c", "c"])
para("注：速度有效值分区参考 ISO 10816-3 / GB/T 6075 中小型机械（Class I）振动烈度分区的"
     "A/B、B/C、C/D 边界值 0.71 / 1.8 / 4.5 mm/s，用于定性参考；"
     "实际限值应按被监测设备的类别、功率与安装方式查标准原文确定。", size=9.5, align="justify")

# ================================================================= 五
h("五、标定与误差分析", 1)
h("5.1 标定方法", 2)
bullets([
    "静态标定（重力翻转法）：将传感器六个面依次朝下，利用 ±1 g 重力加速度标定零偏与灵敏度，"
    "计算零偏温漂与轴间正交误差。",
    "动态标定（振动台比对法）：将监测器与标准加速度计背靠背刚性安装于振动台，"
    "在 20 / 50 / 100 / 200 / 500 / 1000 Hz 各取 3 个幅值点（0.5 g / 2 g / 5 g）测量，"
    "以标准计读数为参考，用最小二乘法拟合灵敏度系数与幅值线性度，并绘制幅频特性曲线。",
    "标定结果以系数形式写入 Flash，软件计算时调用，更换传感器后可重新标定。",
])
h("5.2 误差来源与对策", 2)
table("表 6　主要误差来源与抑制措施", ["误差来源", "影响", "抑制措施"],
      [["安装方式与安装共振", "1 kHz 以上幅值衰减或放大，误差可达 10 % 以上",
        "采用螺栓/粘接刚性安装；避免磁座；在标定中记录安装条件并做频响修正"],
       ["传感器横向灵敏度", "非测量方向振动串入，典型 2 % ~ 5 %",
        "选低横向灵敏度器件；保证安装面与测量方向垂直"],
       ["ADC 量化与基准噪声", "12 bit、3.3 V 基准下 0.8 mV/LSB，小信号信噪比不足",
        "扩展通道使用 3.0 V 精密基准（如 REF3033）；两级放大使信号占满量程的 50 % 以上"],
       ["频谱泄漏与栅栏效应", "幅值与频率读数偏差",
        "Hanning 窗 + 幅值恢复；必要时做比值校正（插值）细化谱峰频率"],
       ["频率分辨率不足", "特征频率识别偏差最大 Δf/2",
        "轴承类特征采用 4096 点长帧或包络谱细化"],
       ["温度漂移", "零偏随温度漂移（约 0.4 mg/℃ 量级），低速测量误差增大",
        "上电预热 1 min；软件去趋势；必要时加温度补偿表"],
       ["电源与地噪声", "开关电源纹波耦合进入模拟通道",
        "模拟与数字分区布线，单点接地，LDO 后加 π 型滤波，必要时加屏蔽"]],
      [Cm(3.4), Cm(5.4), Cm(6.6)], size=9, align_map=["l", "l", "l"])
h("5.3 指标核算", 2)
bullets([
    "频率测量误差：≤ Δf/2 = 0.78 Hz（N = 2048 / 4096），满足特征频率识别要求。",
    "幅值测量误差：目标 ≤ 5 %（满量程的 50 % 以上使用时），含量化、泄漏与标定残差。",
    "响应时间：采样 2048 点耗时 0.64 s，加上 3 帧确认与计算时间，总报警响应约 0.7 ~ 1.0 s，"
    "满足 ≤ 1 s 指标。",
])

# ================================================================= 六
h("六、测试与验证方案", 1)
h("6.1 单元测试", 2)
bullets([
    "SPI 通信测试：读取 ADXL345 器件 ID（0xE5）验证接口正确；静止时三轴读数应接近 0 g（水平）"
    "或 ±1 g（竖直）。",
    "采样率测试：由定时器产生标准方波自检信号，统计 1 s 内的采样点数，误差应小于 0.1 %。",
    "FFT 正确性测试：由信号源注入已知频率与幅值的正弦信号，检查谱峰位置与幅值读数是否与理论值一致。",
])
h("6.2 系统与功能测试", 2)
bullets([
    "幅频特性测试：振动台扫频 5 Hz ~ 1.6 kHz，绘制幅频响应，检查 −3 dB 带宽与带内平坦度。",
    "故障模拟试验：在转子试验台上模拟转子不平衡（加配重）、联轴器不对中（垫片调整）、"
    "轴承点蚀（预制缺陷轴承）、地脚松动四类典型故障，观察特征值变化规律与报警是否正确。",
    "通信与存储测试：连续运行 2 h，检查数据文件完整性与上位机曲线连续性，统计丢包率。",
])
h("6.3 可靠性测试", 2)
bullets([
    "重复性：同一状态下连续测量 10 次，特征值变异系数应 ≤ 3 %。",
    "一致性：与标准测振仪同点对比，速度有效值偏差应 ≤ 10 %。",
    "续航与温升：满电连续运行，记录续航时间与主控、LDO 温升（应 ≤ 25 K）。",
])
table("表 7　测试项目与验收判据", ["测试项目", "测试条件", "验收判据"],
      [["频率测量准确度", "信号源注入 20 / 100 / 500 / 1000 Hz 正弦", "误差 ≤ 1 % 或 ≤ Δf"],
       ["幅值测量准确度", "振动台比对，1 g ~ 5 g", "误差 ≤ 5 %"],
       ["−3 dB 带宽", "振动台扫频", "覆盖 5 Hz ~ 1.6 kHz"],
       ["故障识别能力", "轴承点蚀等四类故障模拟", "至少正确识别 3 类并给出正确等级"],
       ["报警响应时间", "人为产生越限冲击信号", "≤ 1 s"],
       ["续航时间", "满电、WiFi 每 10 s 上报一次", "≥ 8 h"],
       ["重复性", "同一状态重复 10 次", "变异系数 ≤ 3 %"],
       ["连续运行稳定性", "连续运行 2 h", "无死机、无丢帧、文件完整"]],
      [Cm(3.6), Cm(6.0), Cm(5.8)], size=9, align_map=["l", "l", "l"])

# ================================================================= 七
h("七、进度计划与分工", 1)
para("按 8 周课程设计周期安排，小组按 3 人配置（A 负责硬件、B 负责算法、C 负责交互与通信），"
     "实际人数不同时按列合并或拆分。", align="justify")
table("表 8　进度计划与分工", ["周次", "工作内容", "交付物", "负责人"],
      [["第 1 周", "选题分析、方案论证与器件选型、采购下单", "设计方案、选型清单", "全员"],
       ["第 2 周", "硬件原理图设计、PCB 绘制（或模块化搭建方案）", "原理图、PCB 文件", "A + B"],
       ["第 3 周", "焊接调试；主控最小系统、OLED、传感器通信打通", "可运行的最小系统", "A"],
       ["第 4 周", "采集驱动、FFT 与特征提取算法实现与验证", "算法代码、验证数据", "B"],
       ["第 5 周", "状态判别、阈值自学习、OLED 界面与按键交互", "判别程序、显示界面", "C + B"],
       ["第 6 周", "SD 卡存储、WiFi 通信与上位机软件", "上位机程序", "C"],
       ["第 7 周", "灵敏度标定、系统联调、故障模拟试验与性能测试", "测试报告、标定数据", "全员"],
       ["第 8 周", "设计报告撰写、答辩 PPT、实物演示与视频", "报告、PPT、样机", "全员"]],
      [Cm(1.8), Cm(6.6), Cm(4.4), Cm(2.6)], size=9, align_map=["c", "l", "l", "c"])

# ================================================================= 八
h("八、创新点与预期成果", 1)
h("8.1 方案特点与创新点", 2)
bullets([
    "双通道可扩展架构：数字 MEMS 主通道保证低成本与易实现，IEPE 调理通道预留工业级扩展能力，"
    "同一硬件平台覆盖教学演示与工程应用两种需求。",
    "时域 + 频域 + 包络域三域联合判据：单一 RMS 指标对早期冲击型故障不敏感，"
    "引入峭度与包络谱特征频率幅值后，可提前发现轴承早期缺陷。",
    "基线自学习 + 迟滞消抖：装置能适应不同设备与工况，减少误报，提高现场可用性。",
    "边缘计算：2048 点 FFT 与全部特征计算在本地完成，仅上传特征值，"
    "在 10 s 上报周期下带宽占用不足 1 kbps，适合多点组网。",
])
h("8.2 预期成果", 2)
bullets([
    "震动状态监测器样机 1 台（含外壳、面板与电源）。",
    "嵌入式固件 1 套（驱动 + 算法 + 交互 + 通信，含注释与说明文档）。",
    "上位机监测软件 1 套（实时曲线、报警记录、历史回放）。",
    "设计报告 1 份、答辩 PPT 1 份、实物演示视频 1 段。",
    "测试数据与标定报告 1 份（至少 4 类故障模拟的对比数据）。",
])

# ================================================================= 九
h("九、风险分析与应对措施", 1)
table("表 9　主要风险与应对措施", ["风险项", "可能影响", "应对措施"],
      [["采样率与带宽不足，混叠导致误判", "特征频率识别错误，报警不准",
        "严格联动 ODR 与采样率；加强抗混叠滤波；必要时升级为 ICM-42688-P 或 ADXL355"],
       ["无线通信不稳定或丢包", "上位机曲线中断", "本地 SD 卡先存后传；断点续传队列；必要时改用 RS485 有线"],
       ["安装方式引入高频衰减", "幅值偏小，漏报", "统一采用螺栓/粘接安装；标定记录安装条件并做频响修正"],
       ["算力或内存不足", "帧处理超时，报警延迟", "使用 CMSIS-DSP 优化；必要时降为 1024 点或改用单精度 Q15 定点"],
       ["器件缺货或涨价", "进度延误", "预先确定备选型号（MPU6050/LIS3DH、STM32F411/F103+外部 FPU 方案）"],
       ["振动台等仪器不可用", "无法完成标定", "采用重力翻转法做静态标定；用偏心电机 + 标准测振仪做简易比对"],
       ["阈值设置不当导致误报/漏报", "演示效果差", "基线自学习 + 迟滞消抖；现场可手动调整并记录调整依据"]],
      [Cm(3.8), Cm(4.0), Cm(7.6)], size=9, align_map=["l", "l", "l"])

# ================================================================= 附录
doc.add_page_break()
h("附录 A　主控引脚分配建议", 1)
table("表 10　STM32F407VET6 引脚分配", ["引脚", "功能配置", "连接对象", "说明"],
      [["PA4", "GPIO 输出", "ADXL345 CS", "软件片选，SPI1 主机模式"],
       ["PA5 / PA6 / PA7", "SPI1 SCK / MISO / MOSI", "ADXL345 SCLK / SDO / SDA", "时钟 ≤ 5 MHz"],
       ["PB0", "EXTI0 中断输入", "ADXL345 INT1", "数据就绪 / FIFO 水位中断"],
       ["PA0", "ADC1_IN0", "IEPE 调理输出", "扩展通道，TIM2 触发 + DMA"],
       ["PB6 / PB7", "I2C1 SCL / SDA", "OLED SSD1306", "400 kHz 快速模式"],
       ["PC8 ~ PC12、PD2", "SDIO D0~D3 / CK / CMD", "microSD 卡座", "4 位模式 + FATFS"],
       ["PA2 / PA3", "USART2 TX / RX", "ESP8266-01S", "115200 bps，AT 指令"],
       ["PA9 / PA10", "USART1 TX / RX", "调试串口 / 预留 RS485", "程序调试与现场总线扩展"],
       ["PE0 / PE1", "GPIO 输出", "红灯 / 绿灯（黄灯 PE2）", "串 1 kΩ 限流电阻"],
       ["PB8", "TIM4_CH3 PWM", "蜂鸣器驱动", "经 S8050 三极管驱动"],
       ["PC0 / PC1 / PC2", "GPIO 输入（上拉）", "按键 K1 / K2 / K3", "10 kΩ 上拉 + 100 nF 消抖"],
       ["PA13 / PA14", "SWDIO / SWCLK", "调试器", "SWD 下载与在线调试"],
       ["PC13", "GPIO 输出", "电源指示灯", "串联 2 kΩ 电阻"]],
      [Cm(2.8), Cm(3.8), Cm(4.6), Cm(4.2)], size=9, align_map=["c", "l", "l", "l"])

h("附录 B　任务书差异化调整建议", 1)
bullets([
    [("若要求更高带宽（≥ 5 kHz）：", True),
     ("将主传感器换为 ICM-42688-P（ODR 最高 32 kHz，噪声约 70 µg/√Hz）或 ADXL355，"
      "采样率提升至 12.8 ~ 25.6 kSPS，帧长 4096 点，并在软件中加入多级抽取滤波。", False)],
    [("若要求更高精度（幅值误差 ≤ 1 %）：", True),
     ("扩展通道改用 24 bit Σ-Δ ADC（ADS1256 / ADS131M04）+ 精密基准，"
      "模拟前端增加可编程增益放大器，并做通道间相位一致性标定。", False)],
    [("若要求工业现场总线：", True),
     ("以太网（LAN8720）或 RS485 + Modbus-RTU，采用 4 ~ 20 mA / 隔离电源设计，"
      "外壳改为金属屏蔽并加装接线端子。", False)],
    [("若要求超低功耗/电池长期供电：", True),
     ("关闭 WiFi 改为定时唤醒 + 本地 SD 记录，主控降频至 16 MHz，"
      "OLED 改为段码屏或自动息屏，续航可提升至 40 h 以上。", False)],
    [("若要求多测点组网：", True),
     ("各监测节点以 ESP-NOW 或 LoRa 汇聚至网关，网关统一上报；"
      "节点增加设备号与测点编号，上位机按测点维度展示趋势。", False)],
    [("若要求成本更低（≤ 150 元）：", True),
     ("主控换 STM32F103C8T6 + 定点 FFT（1024 点），显示改 0.91″ OLED，"
      "去掉 IEPE 扩展通道与 SD 卡，仅保留 UART 上报。", False)],
])

doc.save(OUT)
print("saved:", OUT)
