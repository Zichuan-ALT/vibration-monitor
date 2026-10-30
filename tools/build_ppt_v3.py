# -*- coding: utf-8 -*-
"""生成 8.1 振动状态监测器 中期答辩 PPT（完整版 v3）。

定位：既不是纯技术版，也不是原来的 v1，而是「完整但有取舍」：
  · 保留组织性内容——汇报分工、选题背景、进度对照、现场演示、仓库；
  · 保留全部技术内容——架构、选型、信号处理、判定、实现、测试、数据分析；
  · **唯一去掉的是硬件费用**（预算、金额、成本、¥），其余约束（零焊接等）保留。

四人分讲：A 背景与需求 / B 设计与信号处理 / C 实现 / D 测试与数据。
输出：8.1_振动状态监测器_中期答辩_完整版.pptx
"""
import os

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "8.1_振动状态监测器_中期答辩_完整版.pptx")
FIG = os.path.join(ROOT, "figures")
EVD = os.path.join(ROOT, "evidence")
DAT = os.path.join(ROOT, "experiments", "data")

W, H = 13.333, 7.5

C_DARK = RGBColor(0x1F, 0x38, 0x64)
C_ACCENT = RGBColor(0x2E, 0x75, 0xB6)
C_WARN = RGBColor(0xC5, 0x5A, 0x11)
C_GRAY = RGBColor(0x59, 0x59, 0x59)
C_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_GREEN = RGBColor(0x37, 0x8A, 0x3E)

FONT = "Microsoft YaHei"

PARTS = {
    1: ("第一部分", "项目背景与需求", "成员 A", C_ACCENT),
    2: ("第二部分", "方案设计与信号处理", "成员 B", RGBColor(0x8E, 0x44, 0xAD)),
    3: ("第三部分", "实现与服务端", "成员 C", RGBColor(0x1F, 0x8A, 0x8A)),
    4: ("第四部分", "测试与数据分析", "成员 D", C_WARN),
}


# ---------------- 基础工具 ----------------
def set_run(run, size=16, bold=False, color=C_GRAY, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", font)


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width),
                                  Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(4)
    tf.margin_top = tf.margin_bottom = Pt(2)
    return tf


def para(tf, text, size=16, bold=False, color=C_GRAY, space_after=6,
         first=False, align=PP_ALIGN.LEFT, level=0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.level = level
    p.space_after = Pt(space_after)
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color)
    return p


def bullets(slide, left, top, width, height, items, size=15, gap=6):
    tf = textbox(slide, left, top, width, height)
    first = True
    for it in items:
        txt, lv, bd, col = it
        prefix = "" if lv == 0 else "· "
        para(tf, prefix + txt, size=size - (1.5 if lv else 0), bold=bd, color=col,
             space_after=gap, first=first, level=lv)
        first = False
    return tf


def hline(slide, left, top, width, color=C_ACCENT, thickness=1.2):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top),
                                Inches(width), Pt(thickness))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    return sh


def rect(slide, left, top, width, height, color, line=False):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top),
                                Inches(width), Inches(height))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    if not line:
        sh.line.fill.background()
    return sh


def add_picture_fit(slide, path, left, top, max_w, max_h):
    iw, ih = Image.open(path).size
    scale = min(max_w / iw, max_h / ih)
    w, h = iw * scale, ih * scale
    x = left + (max_w - w) / 2
    y = top + (max_h - h) / 2
    return slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))


def crop_bottom(src, pad=18):
    import tempfile
    im = Image.open(src).convert("RGB")
    w, h = im.size
    bg = im.getpixel((1, h - 1))
    px = im.load()

    def same(p):
        return (abs(p[0] - bg[0]) < 8 and abs(p[1] - bg[1]) < 8
                and abs(p[2] - bg[2]) < 8)

    last = h - 1
    for y in range(h - 1, 0, -1):
        if any(not same(px[x, y]) for x in range(0, w, 2)):
            last = y
            break
    cut = min(h, last + pad)
    if cut >= h - 12:
        return src
    out = os.path.join(tempfile.gettempdir(), "ppt3_" + os.path.basename(src))
    im.crop((0, 0, w, cut)).save(out)
    return out


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, part, title, subtitle=None):
    label, part_name, speaker, color = PARTS[part]
    rect(slide, 0, 0, W, 0.09, color)
    chip = rect(slide, 0.5, 0.34, 2.95, 0.42, color)
    tf = chip.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Pt(2)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "%s · %s" % (label, speaker)
    set_run(r, 12.5, True, C_WHITE)

    tf = textbox(slide, 3.6, 0.26, 9.2, 0.62, MSO_ANCHOR.MIDDLE)
    para(tf, title, 23, True, C_DARK, 0, first=True)
    if subtitle:
        tf2 = textbox(slide, 0.5, 0.86, 12.3, 0.34)
        para(tf2, subtitle, 13, False, C_GRAY, 0, first=True)
        hline(slide, 0.5, 1.26, 12.33, color)
        return 1.42
    hline(slide, 0.5, 0.98, 12.33, color)
    return 1.14


def pageno(slide, n):
    tf = textbox(slide, 12.0, 6.95, 0.85, 0.3)
    para(tf, str(n), 10, False, RGBColor(0xA6, 0xA6, 0xA6), 0,
         first=True, align=PP_ALIGN.RIGHT)


def table(slide, rows, left, top, width, col_ratios, size=12,
          header_fill=C_DARK, row_h=0.36, header_size=12):
    nrow, ncol = len(rows), len(rows[0])
    shape = slide.shapes.add_table(nrow, ncol, Inches(left), Inches(top),
                                   Inches(width), Inches(row_h * nrow))
    tbl = shape.table
    for i, ratio in enumerate(col_ratios):
        tbl.columns[i].width = Inches(width * ratio)
    for ri, row in enumerate(rows):
        tbl.rows[ri].height = Inches(row_h)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = cell.margin_right = Pt(5)
            cell.margin_top = cell.margin_bottom = Pt(2)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(val)
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
                set_run(r, header_size, True, C_WHITE)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE if ri % 2 else C_LIGHT
                set_run(r, size, False, C_GRAY)
    return tbl


def code_box(slide, lines, left, top, width, height, size=11):
    rect(slide, left, top, width, height, RGBColor(0xF6, 0xF8, 0xFA))
    tf = textbox(slide, left + 0.12, top + 0.08, width - 0.24, height - 0.16)
    first = True
    for ln in lines:
        para(tf, ln, size, False, RGBColor(0x33, 0x33, 0x33), 3, first=first)
        first = False


# ---------------- 各页 ----------------
def s_cover(prs):
    s = blank(prs)
    rect(s, 0, 0, W, 2.05, C_DARK)
    tf = textbox(s, 0.9, 0.42, 11.6, 0.4)
    para(tf, "利莫瑞克大学 × LICHPU　工业评估（15 ECTS）· 智慧校园生活实验室 · 主题 8",
         13, False, RGBColor(0xBF, 0xD4, 0xEE), 0, first=True)
    tf = textbox(s, 0.9, 0.86, 11.6, 0.75)
    para(tf, "项目 8.1　振动状态监测器", 36, True, C_WHITE, 0, first=True)
    tf = textbox(s, 0.9, 1.6, 11.6, 0.4)
    para(tf, "中期答辩", 17, False, RGBColor(0xBF, 0xD4, 0xEE), 0, first=True)

    tf = textbox(s, 0.9, 2.55, 11.6, 2.4)
    para(tf, "四部分，由四位成员分别主讲：", 15, False, C_GRAY, 12, first=True)
    for i, (label, name, spk, color) in PARTS.items():
        p = tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = "%s　%s　——　%s" % (label, name, spk)
        set_run(r, 17, True, color)

    hline(s, 0.9, 5.42, 11.5, C_ACCENT, 1.5)
    tf = textbox(s, 0.9, 5.55, 11.6, 0.8)
    para(tf, "小组成员：成员 A ／ 成员 B ／ 成员 C ／ 成员 D", 14, False, C_GRAY, 4,
         first=True)
    para(tf, "答辩日期：2026 年 11 月", 14, False, C_GRAY, 0)
    notes(s, "开场：今天汇报四部分，每人讲一段。第一部分是背景与需求，第二部分设计与信号处理，"
             "第三部分实现与服务端，第四部分测试与数据。")


def s_division(prs, n):
    s = blank(prs)
    top = header(s, 1, "汇报分工与时间分配", "每人约 3 分钟，合计约 12 分钟，留 3 分钟答疑")
    table(s, [
        ["部分", "内容", "页数", "主讲", "负责的实验与材料"],
        ["第一部分", "项目背景、课题约束、需求规格、进度与下一步", "3 页", "成员 A",
         "用户调研、需求与验收标准、责任工程初筛"],
        ["第二部分", "系统架构、关键选型、采样与判定设计", "5 页", "成员 B",
         "架构设计、信号处理推导、判定策略"],
        ["第三部分", "硬件搭建、接口、固件与服务端", "5 页", "成员 C",
         "器件接入、固件开发、MQTT 与仪表板联调"],
        ["第四部分", "传感器标定、两类故障实验、数据分析", "4 页", "成员 D",
         "标定、不平衡与堵塞工况、频谱分析"],
        ["演示环节", "现场演示加配重触发报警并查看历史", "—", "成员 C 操作、成员 D 讲解",
         "答辩当天现场执行"],
    ], 0.55, top + 0.1, 12.2, [0.11, 0.34, 0.09, 0.16, 0.30], size=11.5, row_h=0.55)

    tf = textbox(s, 0.55, top + 4.1, 12.2, 1.0)
    para(tf, "每个部分都由对应成员主讲，其余成员负责该部分的实验与资料准备；"
             "答辩时的提问由主讲人先答，涉及具体数据时由负责该实验的成员补充。",
         12.5, False, C_GRAY, 0, first=True)
    pageno(s, n)
    notes(s, "先说一下分工：四部分四个人讲，各约三分钟。演示环节由成员C 操作、成员D 讲解。")


def s_background(prs, n):
    s = blank(prs)
    top = header(s, 1, "选题背景：为什么要做振动监测")
    bullets(s, 0.55, top + 0.05, 6.35, 4.8, [
        ("旋转机械的常见故障", 0, True, C_DARK),
        ("风机、水泵、电机等设备，故障多由转子不平衡、轴承磨损、地脚松动引起", 1, False, C_GRAY),
        ("这些故障在早期都会表现为振动与温度特征的变化", 1, False, C_GRAY),
        ("传统做法的问题", 0, True, C_DARK),
        ("定期人工巡检：依赖经验，两次巡检之间的突发故障覆盖不到", 1, False, C_GRAY),
        ("事后维修：故障发生后才处理，停机损失与抢修工作量都大", 1, False, C_GRAY),
        ("状态监测的价值", 0, True, C_DARK),
        ("在征兆出现的早期给出提示，把「事后维修」变成「预测性维护」", 1, False, C_GRAY),
    ], size=14, gap=5)

    rect(s, 7.15, top + 0.1, 5.65, 4.5, C_LIGHT)
    tf = textbox(s, 7.4, top + 0.3, 5.2, 4.1)
    para(tf, "本课题的约束条件", 16, True, C_DARK, 10, first=True)
    for t in ["零焊接：全部选用已焊针的成品模块，杜邦线直插，不需要电烙铁",
              "器件公开可得、资料齐全，可在两校复制",
              "必须有两路以上有意义的数据输入，能区分不同故障",
              "数据与判定都在本地完成，断网时仍可用"]:
        p = tf.add_paragraph()
        p.space_after = Pt(9)
        r = p.add_run()
        r.text = "· " + t
        set_run(r, 13.5, False, C_GRAY)
    pageno(s, n)
    notes(s, "背景很简单：这类设备在坏之前，振动和温度先变。传统靠人听靠人摸，"
             "发现时往往已经晚了。所以要做一个小型的、能持续看的状态监测。"
             "右边是本课题的约束：零焊接、器件通用、两路以上数据、本地判定。")


def s_requirements(prs, n):
    s = blank(prs)
    top = header(s, 1, "需求规格：功能、非功能与验收标准")
    cards = [
        ("11 条", "功能需求 FR",
         "振动采集与特征提取、温度监测、本地存储、联网上报、实时显示、"
         "历史查询、三级报警、维修建议、参数配置、断网续传", C_ACCENT),
        ("9 条", "非功能需求 NFR",
         "零焊接装配、采样率 ≥ 1 kHz、无线上报、数据不丢失、可用性与可维护性、"
         "电气安全、可复制性等 9 条", RGBColor(0x8E, 0x44, 0xAD)),
        ("13 条", "验收标准 AC",
         "每条均可量化验证：采样率误差 ≤ 0.5%、静止模长 1.000 ± 0.2 g、"
         "报警响应 ≤ 10 s、断网零丢失等", RGBColor(0x1F, 0x8A, 0x8A)),
    ]
    x = 0.55
    for num, name, desc, color in cards:
        rect(s, x, top + 0.08, 3.95, 3.5, C_LIGHT)
        rect(s, x, top + 0.08, 3.95, 0.07, color)
        tf = textbox(s, x + 0.25, top + 0.3, 3.45, 3.1)
        para(tf, num, 30, True, color, 4, first=True)
        para(tf, name, 15, True, C_DARK, 10)
        para(tf, desc, 12.5, False, C_GRAY, 0)
        x += 4.13

    tf = textbox(s, 0.55, top + 3.85, 12.2, 1.2)
    para(tf, "截至中期：13 条验收标准中 11 条已完全达成，2 条在收尾", 15, True, C_DARK,
         6, first=True)
    para(tf, "已达成的关键指标：采样率误差 −0.03%（判据 ≤ 0.5%）、报警响应 6 s（判据 ≤ 10 s）、"
             "断网补发 5/5 条零丢失", 13, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "需求分三类：11 条功能、9 条非功能、13 条验收标准。"
             "验收标准全部可量化，比如采样率误差、静止模长、报警响应时间。"
             "中期看，13 条里 11 条已完全达成。")


def s_arch(prs, n):
    s = blank(prs)
    top = header(s, 2, "系统架构：三层结构，判定在边缘完成")
    layers = [
        ("传感层", "MPU6050（三轴加速度，I²C）＋ DS18B20（单总线温度）", C_ACCENT),
        ("边缘层", "ESP32：1 kSPS 采样 → FFT → 带限 RMS → 基线自学习 → 迟滞判定 "
                   "→ SD 本地存储 → MQTT 上报", RGBColor(0x8E, 0x44, 0xAD)),
        ("服务层", "Node-RED：MQTT broker ＋ 仪表板（实时/趋势/历史）＋ SQLite 持久化",
         RGBColor(0x1F, 0x8A, 0x8A)),
    ]
    y = top + 0.1
    for name, desc, color in layers:
        rect(s, 0.55, y, 0.1, 1.0, color)
        tf = textbox(s, 0.8, y, 12.0, 1.0)
        para(tf, name, 15.5, True, color, 3, first=True)
        para(tf, desc, 13, False, C_GRAY, 0)
        y += 1.15

    rect(s, 0.55, top + 3.65, 12.2, 1.55, C_LIGHT)
    tf = textbox(s, 0.8, top + 3.8, 11.7, 1.3)
    para(tf, "关键设计原则", 14.5, True, C_DARK, 4, first=True)
    para(tf, "所有特征提取与状态判定都在 ESP32 上完成，只上传特征值而不上传原始波形。"
             "一帧原始数据约 2 KB，特征报文只有约 180 字节，带宽降到不足十分之一；"
             "更重要的是采样与判定不依赖网络，断网时设备照常测量与判断。",
         12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "架构分三层。最要紧的一点是判定放在边缘端，只上传特征值。"
             "这样带宽省十分之九，断网也不影响监测。")


def s_choice(prs, n):
    s = blank(prs)
    top = header(s, 2, "关键选型：每个决策都有对比与理由")
    table(s, [
        ["决策点", "备选方案", "本方案选择", "主要理由"],
        ["振动传感器", "压电式 / ADXL345 / MPU6050", "MPU6050",
         "I²C 接口简单、量程满足需求、有已焊针版本"],
        ["温度传感器", "NTC / 热电偶 / DS18B20", "DS18B20",
         "数字输出免标定、单总线、防水探头"],
        ["计算位置", "云端处理 / 边缘处理", "边缘处理",
         "带宽、延迟、离线可用、隐私"],
        ["上报协议", "HTTP / MQTT", "MQTT", "轻量、QoS、发布订阅"],
        ["数据存储", "云数据库 / SQLite / SD 卡", "SQLite ＋ SD 卡",
         "SQLite 零配置、SD 卡保证断电不丢"],
    ], 0.55, top + 0.08, 12.2, [0.13, 0.28, 0.18, 0.41], size=11.5, row_h=0.52)

    tf = textbox(s, 0.55, top + 3.55, 12.2, 0.9)
    para(tf, "选型的共同原则：在满足需求的前提下，优先选「接口简单、资料丰富、易复现」的方案，"
             "并保证全程零焊接——所有模块买来就是焊好排针的，用杜邦线与面包板连接即可。",
         13, False, C_GRAY, 0, first=True)
    pageno(s, n)
    notes(s, "五个决策点。共同原则是：接口简单、资料多、容易复现，而且全程零焊接。")


def s_sampling(prs, n):
    s = blank(prs)
    top = header(s, 2, "采样参数是怎么定下来的")
    bullets(s, 0.55, top + 0.05, 6.3, 4.8, [
        ("① 采样率 fs = 1 kHz", 0, True, C_DARK),
        ("分析带宽取 10 ~ 260 Hz；Nyquist 要求 ≥ 520 Hz，取 1 kHz 留约 4 倍余量", 1, False, C_GRAY),
        ("同时让 MPU6050 内部 DLPF（256 Hz）的过渡带落在有效带宽之外", 1, False, C_GRAY),
        ("② 帧长 N = 1024", 0, True, C_DARK),
        ("Δf = fs / N = 0.977 Hz；帧长 1.024 s", 1, False, C_GRAY),
        ("帧再长 → 分辨率更高但报警迟钝；帧再短 → 1× 与相邻谐波分不开", 1, False, C_GRAY),
        ("实测转频 34 Hz，相邻谐波间隔也是 34 Hz，0.977 Hz 足以区分", 1, False, C_GRAY),
        ("③ 量程 ±4 g", 0, True, C_DARK),
        ("0.7 g 配重实测峰值 3.2 g：取 ±2 g 会削顶，取 ±8 g 则分辨率减半", 1, False, C_GRAY),
    ], size=13.5, gap=4)

    code_box(s, [
        "// 采样定时：绝对时刻对齐，误差不随时间累积",
        "uint32_t next = micros();",
        "for (uint16_t i = 0; i < N; i++) {",
        "  while ((int32_t)(micros() - next) < 0) { }",
        "  next += 1000;                 // 严格 1 ms 步进",
        "  mpu.getAcceleration(&ax, &ay, &az);",
        "}",
        "",
        "// 实测：1024 点耗时 1023 ms → 999.7 Hz（−0.03 %）",
    ], 7.05, top + 0.15, 5.75, 2.95, size=10.5)

    tf = textbox(s, 7.05, top + 3.3, 5.75, 1.5)
    para(tf, "为什么不用 delay(1)", 14, True, C_DARK, 4, first=True)
    para(tf, "delay() 每次都会叠加调度误差，累积 1024 次后可达数十毫秒；用 micros() "
             "对齐绝对时刻，只等「该采的时刻」，累计误差不增长。",
         12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "三个参数的推导：采样率按 4 倍 Nyquist；帧长在分辨率与实时性之间折中；"
             "量程按实测峰值留余量。定时用绝对时刻对齐，这是精度来源。")


def s_rms(prs, n):
    s = blank(prs)
    top = header(s, 2, "为什么必须带限计算 RMS", "这是本项目在实现中修正的一个设计")
    bullets(s, 0.55, top + 0.05, 6.35, 4.8, [
        ("现象", 0, True, C_WARN),
        ("模块静止时，时域 RMS 仍有 0.05 g；按噪声密度推算应只有 0.006 g，相差约 10 倍",
         1, False, C_GRAY),
        ("原因", 0, True, C_WARN),
        ("设备并非绝对静止——桌面因人员走动、建筑振动而轻微运动", 1, False, C_GRAY),
        ("这些 < 10 Hz 的低频成分与「设备自身振动」是两回事，却被算进了 RMS", 1, False, C_GRAY),
        ("修正", 0, True, C_GREEN),
        ("先 FFT，再在幅度谱上按 Parseval 定理只累加 10 ~ 260 Hz 的能量", 1, False, C_GRAY),
        ("静止读数 0.0053 g（与手册一致）；运转读数 0.0574 g", 1, True, C_GREEN),
        ("结论", 0, True, C_DARK),
        ("「看起来最朴素的 RMS」也必须先明确要统计什么频段", 1, True, C_DARK),
    ], size=13.5, gap=4)

    code_box(s, [
        "RMS_band = sqrt( Σ_{f∈[10,260]Hz} A_k² / 2 )",
        "",
        "A_k：Hann 窗修正后的单边幅度谱峰值",
        "Hann 窗相干增益 0.5 → 单边幅度需乘 2 修正",
        "",
        "对比（模块静置）：",
        "  时域 RMS（全带）      0.050 g",
        "  带限 RMS（10~260 Hz） 0.0053 g   ← 与手册一致",
        "  运转时带限 RMS        0.0574 g",
    ], 7.1, top + 0.25, 5.7, 3.2, size=10.5)

    tf = textbox(s, 7.1, top + 3.65, 5.7, 1.4)
    para(tf, "对判定的影响", 14, True, C_DARK, 4, first=True)
    para(tf, "若用全带 RMS，噪声底被抬高，正常/注意/报警三档分不开；带限后静止与运转的区分度反而更高。",
         12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "这一页是信号处理里最关键的一处修正。直接用整段时域求 RMS，静止时居然有 0.05 g，"
             "跟手册差 10 倍。查下来是桌面在动。改成频域带限积分就对了。")


def s_decision(prs, n):
    s = blank(prs)
    top = header(s, 2, "状态判定：自学习基线 + 迟滞 + 下限保护")
    bullets(s, 0.55, top + 0.05, 6.3, 4.8, [
        ("为什么不用固定阈值", 0, True, C_WARN),
        ("实测：同一台风扇在不同安装状态下，振动水平相差 3.7 倍", 1, False, C_GRAY),
        ("任何固定阈值都必然在某些场合误报或漏报", 1, False, C_GRAY),
        ("因此改为基线自学习", 0, True, C_DARK),
        ("上电采 15 帧（约 16 s）→ 得 μ、σ", 1, False, C_GRAY),
        ("注意阈值 = μ + 3σ；报警阈值 = μ + 6σ", 1, False, C_GRAY),
        ("迟滞：升级需连续 3 帧确认，降级需 5 帧（约 6 s / 10 s）", 1, False, C_GRAY),
        ("下限保护（来自一次真实教训）", 0, True, C_GREEN),
        ("自学习时若风扇恰在停转，学到的 μ 仅 0.0053 g，推算阈值约 0.008 g → 一启动就误报",
         1, False, C_GRAY),
        ("故规定：μ < 0.010 g 时拒绝该基线，状态置 uncalibrated", 1, True, C_GREEN),
    ], size=13.5, gap=3)

    table(s, [
        ["自学习结果", "μ", "σ", "注意阈值", "报警阈值"],
        ["第一次（固定较紧）", "0.0526 g", "0.0079 g", "0.0763 g", "0.1000 g"],
        ["第二次（支撑偏软）", "0.0942 g", "0.0112 g", "0.1279 g", "0.1676 g"],
    ], 7.15, top + 0.25, 5.65, [0.30, 0.17, 0.17, 0.18, 0.18], size=11.5, row_h=0.42)

    tf = textbox(s, 7.15, top + 1.85, 5.65, 2.7)
    para(tf, "自学习确实起了自适应作用", 14, True, C_DARK, 5, first=True)
    para(tf, "两次 μ 相差近 1.8 倍，阈值随之上移。若用固定阈值 0.100 g，"
             "第二场的正常运转就会被判为报警。", 12.5, False, C_GRAY, 7)
    para(tf, "判定效果", 14, True, C_DARK, 3)
    para(tf, "加 0.35 g 配重后 RMS 1.47 g（报警阈值的 14.7 倍），0.70 g 时 2.68 g（26.8 倍），"
             "三档被干净分开，迟滞无抖动。", 12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "阈值不写死。同一台风扇换个安装状态就差 3.7 倍。所以用自学习，并加了保护："
             "学到的基线太小就拒绝，防止停机时学到错误基线。")


def s_hardware(prs, n):
    s = blank(prs)
    top = header(s, 3, "硬件实现：完整样机与装配方式")
    add_picture_fit(s, os.path.join(EVD, "setup_overview.jpg"), 0.5, top + 0.05,
                    3.9, 4.5)
    add_picture_fit(s, os.path.join(EVD, "mpu6050_mount.jpg"), 4.55, top + 0.05,
                    3.9, 4.5)
    tf = textbox(s, 0.5, top + 4.62, 3.9, 0.35)
    para(tf, "整体装置：风机＋传感器＋主控＋面包板", 11, False, C_GRAY, 0, first=True,
         align=PP_ALIGN.CENTER)
    tf = textbox(s, 4.55, top + 4.62, 3.9, 0.35)
    para(tf, "传感器安装：胶带刚性固定于风扇外框", 11, False, C_GRAY, 0, first=True,
         align=PP_ALIGN.CENTER)

    tf = textbox(s, 8.75, top + 0.1, 4.1, 4.6)
    para(tf, "装配方式", 15, True, C_DARK, 6, first=True)
    para(tf, "全部模块选用已焊针版本，用杜邦线直插面包板连接，全程无焊接。", 12.5,
         False, C_GRAY, 10)
    para(tf, "三个模块的位置与理由", 15, True, C_DARK, 6)
    for t in ["MPU6050 贴在风扇外框边角：位于扇叶旋转圆之外不干涉，同时与外框保持刚性连接",
              "DS18B20 探头置于进风面前方、电机轮毂边缘：对风量变化敏感",
              "ESP32 与 SD 模块放在面包板一侧：避免杜邦线跨越扇叶"]:
        p = tf.add_paragraph()
        p.space_after = Pt(6)
        r = p.add_run()
        r.text = "· " + t
        set_run(r, 12, False, C_GRAY)
    pageno(s, n)
    notes(s, "左边是整个装置，右边是传感器怎么固定的。所有模块都是买来就焊好排针的，"
             "用杜邦线直插，全程没有焊接。传感器贴在外框边角，避开扇叶。")


def s_pins(prs, n):
    s = blank(prs)
    top = header(s, 3, "硬件接口与引脚分配")
    table(s, [
        ["外设", "接口", "信号线", "速率 / 分辨率", "备注"],
        ["MPU6050", "I²C 400 kHz", "SDA=21, SCL=22", "1 ms / 8192 LSB/g", "地址 0x68（AD0 悬空）"],
        ["DS18B20", "1-Wire", "DQ=5", "2 s / 0.0625 ℃", "板载 4.7 kΩ 上拉"],
        ["microSD", "SPI", "18 / 19 / 23 / 13", "每帧一行 CSV", "VCC 必须 5 V"],
        ["MQTT", "TCP 1883", "Wi-Fi", "2 s / 约 180 B", "QoS 1"],
    ], 0.55, top + 0.06, 12.2, [0.15, 0.16, 0.22, 0.23, 0.24], size=12, row_h=0.42)

    tf = textbox(s, 0.55, top + 2.05, 5.6, 2.3)
    para(tf, "接口设计要点", 14.5, True, C_DARK, 5, first=True)
    para(tf, "三组接口互不冲突：I²C 占 21/22，单总线占 5，SPI 占 18/19/23/13。",
         12.5, False, C_GRAY, 5)
    para(tf, "SD 模块的 VCC 必须接 5 V：板载稳压芯片压差不足，接 3.3 V 时模块不工作"
             "（这是实测踩到的坑）。", 12.5, False, C_GRAY, 0)
    add_picture_fit(s, os.path.join(FIG, "figS_wiring.png"), 6.3, top + 1.95, 6.45, 2.9)
    pageno(s, n)
    notes(s, "接口这一页。三组总线互不冲突。要特别提醒：SD 模块必须接 5 V，接 3.3 V 不工作。")


def s_firmware(prs, n):
    s = blank(prs)
    top = header(s, 3, "固件结构与三个可靠性机制")
    bullets(s, 0.55, top + 0.05, 6.3, 1.9, [
        ("主循环结构：先测量，后联网", 0, True, C_DARK),
        ("每轮先完成采样 → FFT → 特征 → 判定，存入环形缓冲，再处理网络", 1, False, C_GRAY),
        ("即使 MQTT 重连阻塞数秒，也不会漏掉本周期的采样点", 1, False, C_GRAY),
    ], size=13.5, gap=4)

    table(s, [
        ["机制", "实现", "解决的问题"],
        ["环形缓冲 + 补发", "60 条环形缓冲（约 2 min），恢复后按序补发", "网络中断期间数据不丢"],
        ["NTP 时间戳", "报文 ts 取测量时刻，不取发送时刻", "补发数据仍落在正确的时间位置"],
        ["SD 本地记录", "每帧追加一行并 flush；报警瞬间另存 1024 点波形",
         "掉电最多丢 1 条；报警可事后做频谱诊断"],
        ["基线下限保护", "自学习 μ < 0.010 g 时拒绝基线，置 uncalibrated",
         "停机时学到错误基线导致误报"],
    ], 0.55, top + 2.05, 12.2, [0.17, 0.46, 0.37], size=11.5, row_h=0.5)

    tf = textbox(s, 0.55, top + 4.45, 12.2, 0.6)
    para(tf, "四个机制互相独立：任一环节失效时系统降级而不断链——例如 SD 卡故障只关闭本地记录，"
             "联网上报不受影响。", 12.5, True, C_ACCENT, 0, first=True)
    pageno(s, n)
    notes(s, "固件讲三件事：主循环顺序、四个可靠性机制、降级设计。核心是「先测量后联网」。")


def s_issues(prs, n):
    s = blank(prs)
    top = header(s, 3, "实现中解决的三个技术问题")
    items = [
        ("① I²C 在 1 kHz 采样下偶发丢帧", C_ACCENT,
         ["现象：连续采样 100 帧，偶有无效帧，数据不连续",
          "定位：1 ms 周期内读 14 字节（含温度）超出 I²C 时间开销",
          "处理：I²C 提至 400 kHz，且只读 6 字节加速度（温度另开 2 s 周期）→ 丢帧率降到 0"]),
        ("② arduinoFFT 2.0.4 的 API 与 1.x 教程不同", C_WARN,
         ["现象：照抄 1.x 写法直接编译失败",
          "原因：2.x 改为模板类 ArduinoFFT<T>，方法名全小写",
          "处理：ArduinoFFT<double> FFT(...)；windowing(FFTWindow::Hann, FFTDirection::Forward)"]),
        ("③ 库对 WHO_AM_I 做了归一化", C_GREEN,
         ["现象：getDeviceID() 返回 0x34 而非手册的 0x68，疑似假芯片",
          "原因：该库取 >>1 后的低 6 位，寄存器原值确为 0x68",
          "处理：自检程序同时读 0x75 原始值与库返回值，两者都核对 → 确认原厂器件"]),
    ]
    y = top + 0.1
    for title, color, lines in items:
        rect(s, 0.55, y, 0.08, 1.5, color)
        tf = textbox(s, 0.8, y, 12.0, 1.5)
        para(tf, title, 15, True, color, 4, first=True)
        for ln in lines:
            para(tf, ln, 12.5, False, C_GRAY, 2)
        y += 1.68
    pageno(s, n)
    notes(s, "三个真实遇到过的问题。第二个和第三个属于「文档与实现不一致」，"
             "经验是：关键判据要直读底层寄存器，第三方封装不能全信。")


def s_server(prs, n):
    s = blank(prs)
    top = header(s, 3, "服务端与仪表板：现场报警的实际效果")
    add_picture_fit(s, crop_bottom(os.path.join(EVD, "dashboard_live_alarm.png")),
                    0.5, top + 0.05, 7.5, 4.5)
    tf = textbox(s, 0.5, top + 4.62, 7.5, 0.3)
    para(tf, "加入 0.35 g 配重后的仪表板实况（真实截图）", 11, False, C_GRAY, 0,
         first=True, align=PP_ALIGN.CENTER)

    tf = textbox(s, 8.35, top + 0.1, 4.5, 4.6)
    para(tf, "仪表板实现的功能", 15, True, C_DARK, 6, first=True)
    for t in ["实时：振动 RMS、被测体表面温度",
              "三级状态：正常 / 注意 / 报警",
              "维修建议：按状态给出处理建议",
              "趋势：最近 1 小时 RMS 曲线",
              "历史：SQLite 查询与曲线"]:
        p = tf.add_paragraph()
        p.space_after = Pt(5)
        r = p.add_run()
        r.text = "· " + t
        set_run(r, 12.5, False, C_GRAY)
    para(tf, "服务端分工", 15, True, C_DARK, 6)
    para(tf, "aedes 提供 MQTT broker；Node-RED 订阅后分三路：仪表板控件、"
             "SQLite 入库、日志文件。服务端不做判定——判定在设备端完成。",
         12, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "这是加配重之后的真实截图：表盘整圈变红、状态报警、给维修建议，"
             "右边趋势图能看到 RMS 由 0.06 陡升到 0.6 的阶跃。服务端只做展示和存储，不做判定。")


def s_calib(prs, n):
    s = blank(prs)
    top = header(s, 4, "传感器标定与验证")
    table(s, [
        ["通道", "项目", "方法", "实测结果", "结论"],
        ["振动", "器件身份", "直读 0x75 寄存器原始值", "0x68（库归一化返回 0x34）", "原厂器件"],
        ["振动", "重力模长", "静止读三轴合成模长", "0.991 ~ 1.002 g", "偏差 ≤ 0.9 %"],
        ["振动", "噪声底", "1024 点 × 多帧，带限 RMS", "0.0053 g", "与推算同量级"],
        ["振动", "采样率", "实测帧时长反算", "999.7 Hz（−0.03 %）", "优于判据 16 倍"],
        ["温度", "稳定性", "静置 90 s，连续读数 53 次", "极差 0.25 ℃，噪声 < 1 LSB", "成段恒定"],
        ["温度", "响应性", "手指握持探头 30 s", "+3.25 ℃，τ ≈ 25~30 s", "一阶热响应"],
        ["温度", "恢复性", "松开后 100 s", "−4.50 ℃，无迟滞", "可逆"],
    ], 0.55, top + 0.06, 12.2, [0.08, 0.13, 0.30, 0.28, 0.21], size=11.5, row_h=0.44)

    tf = textbox(s, 0.55, top + 3.6, 12.2, 1.2)
    para(tf, "温度响应曲线的设计含义", 14.5, True, C_DARK, 4, first=True)
    para(tf, "每 4 秒升温量依次为 +0.81、+0.56、+0.50、+0.38、+0.31、+0.25、+0.19 ℃，"
             "速率单调递减 → 典型一阶响应，τ ≈ 25~30 s。结论：温度报警必须设分钟级确认延时，"
             "不能照搬振动通道的秒级迟滞，否则手碰一下探头就会误报。", 12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "标定分两路。振动做四项自洽检查，温度做稳定性、响应、恢复三项。"
             "重点是温度时间常数约 30 秒，这决定了温度报警不能用秒级迟滞。")


def s_imb(prs, n):
    s = blank(prs)
    top = header(s, 4, "故障实验一：转子不平衡", "单片扇叶贴蓝丁胶配重，距轴心约 35 mm")
    table(s, [
        ["工况", "配重", "带内 RMS", "1× 峰值", "峰值频率"],
        ["① 基准", "0 g", "0.0574 g", "0.0292 g", "32.2 Hz"],
        ["② 轻微", "0.35 g", "1.4723 g（+25.7 倍）", "1.6605 g（+56.9 倍）", "33.2 Hz"],
        ["③ 明显", "0.70 g", "2.6816 g（+46.7 倍）", "2.8015 g（+95.9 倍）", "31.2 Hz"],
    ], 0.55, top + 0.04, 6.35, [0.16, 0.16, 0.27, 0.26, 0.15], size=11, row_h=0.4)

    tf = textbox(s, 0.55, top + 1.7, 6.35, 3.0)
    para(tf, "灵敏度", 14, True, C_DARK, 3, first=True)
    para(tf, "0.35 g 仅占转子质量（约 90 g）的 0.4 %，却使 RMS 上升 25.7 倍、1× 峰值上升 "
             "56.9 倍。基准工况的带内 RMS 已是噪声底的约 10 倍，加配重后拉大到约 270 倍，"
             "轻微劣化也在可检测范围内。", 12.5, False, C_GRAY, 6)
    para(tf, "亚线性与共振", 14, True, C_DARK, 3)
    para(tf, "配重加倍只带来 1.82 倍响应。用 F = m·r·ω² 校核：0.35 g 在 35 mm、32 Hz 下约 0.49 N，"
             "对应刚体加速度 0.56 g，而实测峰值 1.66 g，放大倍数 Q ≈ 3.7 —— "
             "系统工作在共振区附近，响应被阻尼压制。", 12.5, False, C_GRAY, 0)

    add_picture_fit(s, os.path.join(DAT, "spectra_compare.png"), 7.0, top + 0.1, 5.75, 4.9)
    pageno(s, n)
    notes(s, "不平衡三级工况，每档 6 到 7 帧。右边频谱三条曲线在 33 Hz 附近依次抬升，"
             "谐波同步上升——这是不平衡的典型谱特征。")


def s_blk(prs, n):
    s = blank(prs)
    top = header(s, 4, "故障实验二：风道堵塞", "硬纸板遮挡进风面，每组稳定 5 min 后采集")
    table(s, [
        ["工况", "遮挡", "表面温度", "振动 RMS", "1× 峰值", "转频"],
        ["① 正常", "0 %", "24.05 ℃", "0.0867 g", "0.0773 g", "34.5 Hz"],
        ["② 中度", "50 %", "24.67 ℃（+0.62）", "0.0589 g（−32 %）", "0.0487 g", "34.2 Hz"],
        ["③ 重度", "80 %", "24.85 ℃（+0.80）", "0.0605 g（−30 %）", "0.0242 g（−69 %）", "31.5 Hz"],
    ], 0.55, top + 0.04, 12.2, [0.09, 0.09, 0.20, 0.20, 0.21, 0.11], size=11.5, row_h=0.42)

    tf = textbox(s, 0.55, top + 1.85, 5.85, 3.4)
    para(tf, "振动反而下降", 14.5, True, C_WARN, 4, first=True)
    para(tf, "风扇振动包含机械与气动两部分。堵塞使风量减少、气动激励减弱，总振动随之下降；"
             "同时电机负载增大导致转速下降（34.5 → 31.5 Hz），散热变差导致温度上升。",
         12.5, False, C_GRAY, 7)
    para(tf, "单以 RMS 作判据不可靠", 14.5, True, C_WARN, 4)
    para(tf, "RMS 在 50 % 与 80 % 之间几乎不变（0.0589 与 0.0605），不随堵塞程度单调变化。"
             "真正单调的是温度、1× 峰值与转频三项。", 12.5, False, C_GRAY, 0)

    add_picture_fit(s, os.path.join(EVD, "blockage_test.jpg"), 6.6, top + 1.85, 6.2, 3.4)
    pageno(s, n)
    notes(s, "堵塞的结论和直觉相反：越堵振动越小，因为气动激励减弱了。所以单看 RMS 会把劣化"
             "误判成「状态改善」。真正单调的是温度、1× 峰值、转频。")


def s_stiff(prs, n):
    s = blank(prs)
    top = header(s, 4, "关键技术发现：支撑刚度的影响达 3.7 倍")
    table(s, [
        ["状态", "1× 峰值", "带内 RMS", "峰值 / RMS", "说明"],
        ["固定良好", "0.042 ~ 0.062 g", "0.0556 g", "≈ 1.0", "能量集中在 1× 附近"],
        ["风扇晃动", "0.066 ~ 0.079 g", "0.1966 g", "0.36", "额外能量分布在非 1× 频段"],
    ], 0.55, top + 0.04, 12.2, [0.15, 0.22, 0.18, 0.15, 0.30], size=11.5, row_h=0.42)

    bullets(s, 0.55, top + 1.65, 6.0, 3.4, [
        ("现象", 0, True, C_WARN),
        ("同一台风扇、同一转速，RMS 出现 0.055 g 与 0.207 g 两种水平", 1, False, C_GRAY),
        ("第一次判断（错误）", 0, True, C_WARN),
        ("按压传感器模块 → RMS 降 57 %，以为胶带松了", 1, False, C_GRAY),
        ("重新粘贴后反而升到 0.207 g —— 按压时手指同时给整个结构加了阻尼", 1, False, C_GRAY),
        ("第二次判断（正确）", 0, True, C_GREEN),
        ("改为按压风扇底座 → RMS 降到 0.0556 g（−73 %）", 1, False, C_GRAY),
        ("结论：不是传感器松了，而是风扇整体在桌面上跳动", 1, True, C_GREEN),
    ], size=12.5, gap=3)

    tf = textbox(s, 6.85, top + 1.65, 5.95, 3.4)
    para(tf, "可用的在线判别指标", 14.5, True, C_DARK, 5, first=True)
    para(tf, "把「1× 峰值 / 带内 RMS」作为辅助诊断量：比值 ≈ 1.0 表示能量集中在转频、测量可信；"
             "显著小于 1 说明叠加了刚体运动，当前数据不可用于判故障。", 12.5, False, C_GRAY, 8)
    para(tf, "工程意义", 14.5, True, C_DARK, 4)
    para(tf, "支撑刚度的影响（3.7 倍）远大于传感器粘接质量。根本原因是传感器无法区分"
             "「设备自身的结构振动」与「设备整体的刚体运动」——这也是工业标准要求"
             "振动测量必须相对于刚性基础进行的原因。", 12.5, False, C_GRAY, 8)
    para(tf, "方法论：怀疑某个因素时，应设计只改变该因素的实验；若两个变量同时被改变，"
             "即便结果符合预期也不能作为证据。", 12.5, True, C_ACCENT, 0)
    pageno(s, n)
    notes(s, "这一页是最有价值的发现，也是一次误判的记录。第一次按压传感器，结论指向「胶松了」，"
             "但重贴后反而更差。第二次改为按压底座才定位到真正原因：风扇在跳。")


def s_progress(prs, n):
    s = blank(prs)
    top = header(s, 1, "进度对照与下一步工作")
    table(s, [
        ["阶段", "原计划", "实际状态", "说明"],
        ["需求与方案", "第 2–3 周", "已完成", "提案与需求规格、实施方案"],
        ["硬件搭建", "第 4–5 周", "已完成", "三模块全部接入并验证（零焊接）"],
        ["固件开发", "第 6–8 周", "已完成", "11 个固件，含 FFT、判定、存储"],
        ["系统集成", "第 9–10 周", "已完成", "MQTT 全链路 + 仪表板 + 数据库"],
        ["测试与实验", "第 11 周", "已完成主要部分", "不平衡三级、堵塞三级、温度标定"],
        ["报告与答辩", "第 12 周", "进行中", "报告已完稿，答辩材料修订中"],
    ], 0.55, top + 0.06, 12.2, [0.16, 0.14, 0.18, 0.52], size=11.5, row_h=0.44)

    tf = textbox(s, 0.55, top + 3.5, 6.0, 2.0)
    para(tf, "整体进度约提前 2 周", 14.5, True, C_GREEN, 5, first=True)
    para(tf, "系统已可完整运行并现场演示，剩余工作集中在报告与细节打磨。",
         12.5, False, C_GRAY, 0)

    tf = textbox(s, 6.85, top + 3.5, 5.95, 2.2)
    para(tf, "下一步工作（按优先级）", 14.5, True, C_DARK, 5, first=True)
    for i, t in enumerate([
        "复合判据写入固件：温度升 + 振动降 → 气动故障",
        "连续运行 24 h 统计误报率",
        "温度绝对标定（冰水混合物 0 ℃）",
        "补充电流通道，覆盖题目要求的三类量",
    ], 1):
        p = tf.add_paragraph()
        p.space_after = Pt(4)
        r = p.add_run()
        r.text = "%d. %s" % (i, t)
        set_run(r, 12.5, False, C_GRAY)
    pageno(s, n)
    notes(s, "进度上整体提前两周，系统已经能跑能演示。剩下的是报告和打磨。"
             "下一步最要紧的是把温度做成复合判据，以及补电流通道。")


def s_demo(prs, n):
    s = blank(prs)
    top = header(s, 1, "现场演示方案（答辩当天）")
    steps = [
        ("①", "打开仪表板，显示「正常」", "RMS 表盘 0.05~0.09 g、温度约 24 ℃", "成员 C"),
        ("②", "确认串口已打印基线 μ、σ 与两个阈值", "检查点：自学习未完成不得进入下一步", "成员 D"),
        ("③", "关闭风扇，单片扇叶贴 0.35 g 蓝丁胶，重新上电", "约 6 s 后表盘变红、状态「报警」",
         "成员 D"),
        ("④", "切到历史页，查看 SQLite 记录中的阶跃", "曲线出现由 0.06 到 1.47 g 的阶跃",
         "成员 A"),
        ("⑤", "断开笔记本 Wi-Fi 30 s 后恢复", "串口显示缓冲 5 条并全部补发", "成员 C"),
    ]
    y = top + 0.1
    for num, act, expect, who in steps:
        rect(s, 0.55, y + 0.06, 0.52, 0.52, C_ACCENT)
        tf = textbox(s, 0.55, y + 0.06, 0.52, 0.52, MSO_ANCHOR.MIDDLE)
        para(tf, num, 18, True, C_WHITE, 0, first=True, align=PP_ALIGN.CENTER)
        tf = textbox(s, 1.25, y, 7.4, 0.75)
        para(tf, act, 13.5, True, C_DARK, 2, first=True)
        para(tf, "预期：" + expect, 12, False, C_GRAY, 0)
        tf = textbox(s, 8.8, y + 0.1, 3.9, 0.5)
        para(tf, who, 12.5, False, C_GRAY, 0, first=True)
        y += 0.95

    tf = textbox(s, 0.55, y + 0.1, 12.2, 0.9)
    para(tf, "失败备用：① 风扇不转——先查供电与接线，1 min 内不能解决则切录屏；"
             "② 配重甩飞——立即断电，换 0.35 g 重压；③ 服务端无响应——重启 node-red，"
             "2 min 未恢复则改为串口 + SD 卡日志的降级演示，并声明这是降级。",
         12, False, C_WARN, 0, first=True)
    pageno(s, n)
    notes(s, "演示五步，约四分钟。第二步是检查点，自学习没完成绝对不能加配重。"
             "下面写了三条失败备用方案。")


def s_end(prs):
    s = blank(prs)
    rect(s, 0, 0, W, 7.5, C_DARK)
    tf = textbox(s, 1.2, 1.9, 11.0, 1.2)
    para(tf, "当前状态", 15, False, RGBColor(0xBF, 0xD4, 0xEE), 10, first=True)
    para(tf, "系统已完整可跑，两类故障可分，四人分工已定", 27, True, C_WHITE, 0)
    hline(s, 1.2, 3.7, 7.5, RGBColor(0xBF, 0xD4, 0xEE), 1.2)

    tf = textbox(s, 1.2, 3.95, 11.0, 1.6)
    para(tf, "项目仓库", 15, False, RGBColor(0xBF, 0xD4, 0xEE), 6, first=True)
    para(tf, "github.com/Zichuan-ALT/vibration-monitor", 17, True, C_WHITE, 4)
    para(tf, "代码、原始数据、实验记录与证据照片均已归档", 13, False,
         RGBColor(0xBF, 0xD4, 0xEE), 0)

    tf = textbox(s, 1.2, 5.75, 11.0, 1.0)
    para(tf, "谢谢，欢迎提问", 24, True, C_WHITE, 6, first=True)
    para(tf, "小组成员：成员 A ／ 成员 B ／ 成员 C ／ 成员 D", 13, False,
         RGBColor(0xBF, 0xD4, 0xEE), 0)
    notes(s, "收尾：系统已经能完整跑起来，两类故障能分开，四个人分工明确。仓库里代码、"
             "原始数据和证据都齐。谢谢，欢迎提问。")


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    s_cover(prs)
    s_division(prs, 2)
    s_background(prs, 3)
    s_requirements(prs, 4)
    s_arch(prs, 5)
    s_choice(prs, 6)
    s_sampling(prs, 7)
    s_rms(prs, 8)
    s_decision(prs, 9)
    s_hardware(prs, 10)
    s_pins(prs, 11)
    s_firmware(prs, 12)
    s_issues(prs, 13)
    s_server(prs, 14)
    s_calib(prs, 15)
    s_imb(prs, 16)
    s_blk(prs, 17)
    s_stiff(prs, 18)
    s_progress(prs, 19)
    s_demo(prs, 20)
    s_end(prs)

    prs.save(OUT)
    print("已生成：", OUT)
    print("页数：", len(prs.slides))


if __name__ == "__main__":
    build()
