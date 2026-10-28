# -*- coding: utf-8 -*-
"""生成 8.1 振动状态监测器 中期答辩 PPT（4 人分讲）。

结构：
  封面 → 汇报分工 → 四部分（每人一部分）→ 结束
每张内容页左上角标注"部分编号 + 主讲人"，便于评委核对四人都有发言。
每张幻灯片均写入备注（演讲要点）。
"""
import os

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "8.1_振动状态监测器_中期答辩.pptx")

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
    2: ("第二部分", "方案设计与选型", "成员 B", RGBColor(0x8E, 0x44, 0xAD)),
    3: ("第三部分", "实现与进度", "成员 C", RGBColor(0x1F, 0x8A, 0x8A)),
    4: ("第四部分", "测试验证与下一步", "成员 D", C_WARN),
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
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
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


def bullets(slide, left, top, width, height, items, size=15, gap=8):
    """items: [(text, level, bold, color)] 或 str"""
    tf = textbox(slide, left, top, width, height)
    first = True
    for it in items:
        if isinstance(it, str):
            txt, lv, bd, col = it, 0, False, C_GRAY
        else:
            txt, lv, bd, col = it
        prefix = "" if lv == 0 else "· "
        para(tf, prefix + txt, size=size - (2 if lv else 0), bold=bd, color=col,
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
    """裁掉图片底部纯背景区域（截图下方常有大片空白），返回可用的路径。"""
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
    out = os.path.join(tempfile.gettempdir(), "ppt_crop_" + os.path.basename(src))
    im.crop((0, 0, w, cut)).save(out)
    return out


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, part, title, subtitle=None):
    """内容页统一页眉：左侧色块标注部分与主讲人。"""
    label, part_name, speaker, color = PARTS[part]
    rect(slide, 0, 0, W, 0.09, color)
    # 部分标签
    chip = rect(slide, 0.5, 0.34, 2.55, 0.42, color)
    tf = chip.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Pt(2)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "%s · %s" % (label, speaker)
    set_run(r, 13, True, C_WHITE)

    # 标题
    tf = textbox(slide, 3.2, 0.26, 9.6, 0.62, MSO_ANCHOR.MIDDLE)
    para(tf, title, 25, True, C_DARK, 0, first=True)
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


# ---------------- 构建 ----------------
def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    n = 0

    # ===== 1 封面 =====
    s = blank(prs)
    n += 1
    rect(s, 0, 0, W, H, C_DARK)
    rect(s, 0, 5.05, W, 0.06, C_ACCENT)
    tf = textbox(s, 1.0, 1.45, 11.3, 0.5)
    para(tf, "工业评估 · Smart Campus Living Lab · 2026/27", 15, False,
         RGBColor(0xA8, 0xC4, 0xE8), 0, first=True)
    tf = textbox(s, 1.0, 2.05, 11.3, 1.0)
    para(tf, "主题 8.1  振动状态监测器", 42, True, C_WHITE, 0, first=True)
    tf = textbox(s, 1.0, 3.15, 11.3, 0.6)
    para(tf, "中期答辩", 26, False, RGBColor(0xA8, 0xC4, 0xE8), 0, first=True)

    tf = textbox(s, 1.0, 5.35, 11.3, 1.2)
    para(tf, "小组成员：成员 A ／ 成员 B ／ 成员 C ／ 成员 D", 17, False, C_WHITE,
         6, first=True)
    para(tf, "答辩日期：2026 年 11 月", 15, False, RGBColor(0xA8, 0xC4, 0xE8), 0)
    notes(s, "开场：各位老师好，我们是第八组，课题是主题 8.1 振动状态监测器。\n"
             "本次汇报我组四人依次发言，分别介绍项目背景与需求、方案设计、"
             "实现进度、测试验证与下一步计划。")

    # ===== 2 汇报分工 =====
    s = blank(prs)
    n += 1
    top = header(s, 1, "汇报分工")
    rows = [["部分", "内容", "主讲"],
            ["第一部分", "项目背景与需求分析", "成员 A"],
            ["第二部分", "方案设计与技术选型", "成员 B"],
            ["第三部分", "实现进度（硬件/固件/服务端）", "成员 C"],
            ["第四部分", "测试验证、关键发现与下一步计划", "成员 D"]]
    table(s, rows, 0.8, top + 0.15, 11.7, [0.18, 0.62, 0.20], size=15,
          row_h=0.52, header_size=15)
    tf = textbox(s, 0.8, 4.55, 11.7, 1.9)
    para(tf, "汇报顺序与时间分配", 15, True, C_DARK, 8, first=True)
    para(tf, "每人约 3 分钟，合计约 12 分钟，留 3 分钟答疑", 14, False, C_GRAY, 6)
    para(tf, "每个部分由一个成员主讲，其余成员负责该部分的实验与资料准备", 14,
         False, C_GRAY, 6)
    para(tf, "演示环节由成员 C 操作设备，成员 D 讲解数据", 14, False, C_GRAY, 0)
    notes(s, "说明分工：四个人各讲一部分，老师可以看到四人都参与了工作。")

    # ===== 3 背景与问题（A） =====
    s = blank(prs)
    n += 1
    top = header(s, 1, "选题背景：为什么要做振动监测")
    bullets(s, 0.8, top + 0.1, 11.7, 4.2, [
        ("旋转机械的常见故障", 0, True, C_DARK),
        ("风机、水泵、电机等设备，故障多由转子不平衡、轴承磨损、地脚松动引起", 1, False, C_GRAY),
        ("这些故障在早期都会表现为振动特征的变化", 1, False, C_GRAY),
        ("传统做法的问题", 0, True, C_DARK),
        ("定期人工巡检：依赖经验，间隔期内故障可能突发", 1, False, C_GRAY),
        ("事后维修：故障发生后才处理，停机损失大", 1, False, C_GRAY),
        ("状态监测的价值", 0, True, C_DARK),
        ("在故障早期发现征兆，把「事后维修」变为「预测性维护」", 1, False, C_GRAY),
        ("本课题的设计目标", 0, True, C_ACCENT),
        ("实现可运行的多参数状态监测样机，所有信号处理与状态判定在设备端完成", 1, False, C_GRAY),
        ("方案可复现，可复制到不同环境（UL ↔ LICHPU）", 1, False, C_GRAY),
        ("系统已可完整运行，具备现场演示条件", 1, False, C_GRAY),
    ], size=15)
    notes(s, "这一部分我讲背景：为什么要做这件事、以及本课题的设计目标。"
             "重点是说明这类故障在早期都有振动特征的变化，因此可以用较简单的方案实现监测。")

    # ===== 4 需求规格（A） =====
    s = blank(prs)
    n += 1
    top = header(s, 1, "需求规格：功能、非功能与验收标准")
    cards = [
        ("11 条", "功能需求 FR",
         "振动采集与特征提取、温度监测、本地存储、联网上报、实时显示、"
         "历史查询、三级报警、维修建议、参数配置、断网续传", C_ACCENT),
        ("9 条", "非功能需求 NFR",
         "采样率 ≥ 1 kHz、无线上报、数据不丢失、可用性与可维护性、"
         "电气安全、可复制性等 9 条", RGBColor(0x8E, 0x44, 0xAD)),
        ("13 条", "验收标准 AC",
         "每条均可量化验证：采样率误差 ≤ 0.5%、静止模长 1.000 ± 0.2 g、"
         "报警响应 ≤ 10 s 等", RGBColor(0x1F, 0x8A, 0x8A)),
    ]
    x = 0.8
    for num, name, desc, color in cards:
        rect(s, x, top + 0.15, 3.78, 3.15, C_LIGHT)
        rect(s, x, top + 0.15, 3.78, 0.08, color)
        tf = textbox(s, x + 0.22, top + 0.32, 3.35, 0.6)
        para(tf, num, 30, True, color, 0, first=True)
        tf = textbox(s, x + 0.22, top + 0.95, 3.35, 0.4)
        para(tf, name, 15, True, C_DARK, 0, first=True)
        tf = textbox(s, x + 0.22, top + 1.42, 3.35, 1.75)
        para(tf, desc, 12, False, C_GRAY, 0, first=True)
        x += 3.96

    tf = textbox(s, 0.8, top + 3.5, 11.7, 1.4)
    para(tf, "截至中期：13 条验收标准中，11 条已完全达成，2 条在收尾", 15, True, C_DARK,
         8, first=True)
    para(tf, "已达成的关键指标：采样率误差 −0.03%（判据 ≤0.5%）、"
             "报警响应 6 s（判据 ≤10 s）", 14, False, C_GRAY, 0)
    notes(s, "需求这一页讲三层：功能 11 条、非功能 9 条、验收 13 条。"
             "重点是验收标准全部可量化，这为后面的测试提供了依据。")

    # ===== 5 系统架构（B） =====
    s = blank(prs)
    n += 1
    top = header(s, 2, "系统架构：三层结构，判定在边缘完成")

    layers = [
        ("传感层", "MPU6050（三轴加速度，I²C）＋ DS18B20（单总线温度）", C_ACCENT),
        ("边缘层", "ESP32：1 kSPS 采样 → FFT → 带限 RMS → 基线自学习 → 迟滞判定 → "
                  "SD 本地存储 → MQTT 上报", RGBColor(0x8E, 0x44, 0xAD)),
        ("服务层", "Node-RED：MQTT broker ＋ 仪表板（实时/趋势/历史）＋ SQLite 持久化",
         RGBColor(0x1F, 0x8A, 0x8A)),
    ]
    y = top + 0.15
    for name, desc, color in layers:
        rect(s, 0.8, y, 1.75, 0.95, color)
        tf = textbox(s, 0.85, y, 1.65, 0.95, MSO_ANCHOR.MIDDLE)
        para(tf, name, 16, True, C_WHITE, 0, first=True, align=PP_ALIGN.CENTER)
        rect(s, 2.55, y, 9.95, 0.95, C_LIGHT)
        tf = textbox(s, 2.75, y, 9.6, 0.95, MSO_ANCHOR.MIDDLE)
        para(tf, desc, 14, False, C_GRAY, 0, first=True)
        y += 1.12

    rect(s, 0.8, top + 3.55, 11.7, 1.5, RGBColor(0xFD, 0xF2, 0xE9))
    tf = textbox(s, 1.0, top + 3.68, 11.3, 1.3)
    para(tf, "关键设计原则", 15, True, C_WARN, 6, first=True)
    para(tf, "所有特征提取与状态判定都在 ESP32 上完成，只上传特征值而不上传原始波形。"
             "这样带宽占用只有几十字节/条，断网时可本地缓存，采样与判定不受网络影响。",
         14, False, C_GRAY, 0)
    notes(s, "架构分三层。最重要的一句话是：判定在边缘完成，不传原始波形。"
             "这既省带宽，也让系统在断网时仍能工作。")

    # ===== 6 关键选型（B） =====
    s = blank(prs)
    n += 1
    top = header(s, 2, "关键选型：每个决策都有对比与理由")
    rows = [["决策点", "备选方案", "本方案选择", "主要理由"],
            ["振动传感器", "压电式 / ADXL345 / MPU6050", "MPU6050",
             "I²C 接口简单、资料丰富、量程满足需求"],
            ["温度传感器", "NTC / 热电偶 / DS18B20", "DS18B20",
             "数字输出免标定、单总线、防水探头"],
            ["计算位置", "云端处理 / 边缘处理", "边缘处理",
             "带宽、延迟、离线可用、隐私"],
            ["上报协议", "HTTP / MQTT", "MQTT", "轻量、QoS、发布订阅"],
            ["数据存储", "云数据库 / SQLite / SD 卡", "SQLite ＋ SD 卡",
             "SQLite 零配置、SD 卡保证断电不丢"]]
    table(s, rows, 0.8, top + 0.15, 11.7, [0.16, 0.30, 0.20, 0.34], size=13,
          row_h=0.66, header_size=13)
    tf = textbox(s, 0.8, top + 4.3, 11.7, 0.8)
    para(tf, "选型的共同原则：在满足需求前提下优先选「接口简单、资料丰富、易于复现」"
             "的方案，便于在不同环境中复制。", 14, False, C_GRAY, 0, first=True)
    notes(s, "这一页是技术设计的核心——不是「我们选了什么」，而是「为什么选它、"
             "为什么不选别的」。老师看的就是这个权衡过程。")

    # ===== 7 信号处理设计（B） =====
    s = blank(prs)
    n += 1
    top = header(s, 2, "信号处理设计：采样参数是如何定下来的")

    tf = textbox(s, 0.8, top + 0.05, 6.1, 4.6)
    para(tf, "采样参数推导", 16, True, C_DARK, 8, first=True)
    para(tf, "关注频段 10~260 Hz → 按 Nyquist 取 fs = 1 kHz（4 倍余量）", 14, False, C_GRAY, 5)
    para(tf, "N = 1024 → 频率分辨率 Δf = fs/N = 0.977 Hz", 14, False, C_GRAY, 5)
    para(tf, "帧长 1.024 s，兼顾实时性与分辨率", 14, False, C_GRAY, 5)
    para(tf, "量程 ±4 g（8192 LSB/g），兼顾分辨率与不平衡工况峰值", 14, False, C_GRAY, 12)

    para(tf, "为什么必须「带限」计算 RMS", 16, True, C_DARK, 8)
    para(tf, "全带 RMS 会把桌面晃动、手持抖动（<10 Hz）算进去", 14, False, C_GRAY, 5)
    para(tf, "实测对比：全带 0.05 g vs 带限 0.005 g，相差 10 倍", 14, False, C_WARN, 5)
    para(tf, "做法：由单边幅度谱按 Parseval 定理求 Σ A²/2", 14, False, C_GRAY, 0)

    rect(s, 7.2, top + 0.1, 5.35, 4.45, C_LIGHT)
    tf = textbox(s, 7.45, top + 0.28, 4.9, 4.1)
    para(tf, "状态判定策略", 16, True, C_DARK, 8, first=True)
    para(tf, "基线自学习：开机采 15 帧，得 μ、σ", 14, False, C_GRAY, 5)
    para(tf, "阈值：注意 = μ+3σ，报警 = μ+6σ", 14, False, C_GRAY, 5)
    para(tf, "迟滞：升级需连续 3 帧确认，降级需 5 帧", 14, False, C_GRAY, 5)
    para(tf, "基线下限保护：μ < 0.010 g 时拒绝该基线", 14, False, C_GRAY, 10)
    para(tf, "为什么用自学习而不是固定阈值", 15, True, C_DARK, 6)
    para(tf, "同一台设备在不同安装状态下的振动水平可差 3 倍以上，"
             "固定阈值必然误报或漏报。", 14, False, C_GRAY, 0)
    notes(s, "这一页讲两个「为什么」：为什么 1 kHz / 1024 点，为什么 RMS 必须带限。"
             "带限那条是我们实测出来的——全带算出来的 RMS 是带限的 10 倍。")

    # ===== 8 硬件实现（C） =====
    s = blank(prs)
    n += 1
    top = header(s, 3, "硬件实现：完整样机与安装方式")
    add_picture_fit(s, os.path.join(ROOT, "evidence", "setup_overview.jpg"),
                    0.75, top + 0.05, 5.4, 3.85)
    add_picture_fit(s, os.path.join(ROOT, "evidence", "mpu6050_mount.jpg"),
                    6.4, top + 0.05, 6.2, 3.85)
    tf = textbox(s, 0.75, top + 3.95, 5.4, 0.32)
    para(tf, "左图：整体装置（风扇＋传感器＋主控＋面包板）", 12, False, C_GRAY, 0,
         first=True, align=PP_ALIGN.CENTER)
    tf = textbox(s, 6.4, top + 3.95, 6.2, 0.32)
    para(tf, "右图：振动传感器安装（胶带刚性固定于风扇外框）", 12, False, C_GRAY, 0,
         first=True, align=PP_ALIGN.CENTER)
    rect(s, 0.75, top + 4.42, 11.85, 0.62, RGBColor(0xE8, 0xF5, 0xE9))
    tf = textbox(s, 0.95, top + 4.5, 11.5, 0.48, MSO_ANCHOR.MIDDLE)
    para(tf, "三部分模块通过标准化接口连接，结构清晰、便于拆装与更换。",
         14, True, C_GREEN, 0, first=True)
    notes(s, "硬件这页看两张照片：左边是整个系统，右边是传感器怎么固定在风扇上。"
             "强调模块通过标准接口连接，方便拆装和替换。")

    # ===== 9 固件演进（C） =====
    s = blank(prs)
    n += 1
    top = header(s, 3, "固件演进：11 个程序，逐个模块验证后再集成")
    rows = [["序号", "固件", "验证内容"],
            ["01–04", "点灯 / 扫 Wi-Fi / 连 Wi-Fi / 发 MQTT", "工具链与网络链路"],
            ["05–06", "I²C 总线扫描 / MPU6050 自检", "传感器是否存在、是否原厂"],
            ["07", "1 kHz 采样 + FFT", "采样率精度、噪声底"],
            ["08", "台架测量（三轴选轴 + 带限 RMS）", "能否测出风扇转频"],
            ["09", "完整节点（双传感 + SD + 上报）", "系统集成"],
            ["10–11", "DS18B20 测试 / SD 卡诊断", "新增模块单独验证"]]
    table(s, rows, 0.8, top + 0.12, 11.7, [0.12, 0.46, 0.42], size=13, row_h=0.55,
          header_size=13)
    rect(s, 0.8, top + 4.0, 11.7, 1.15, RGBColor(0xE8, 0xF1, 0xFA))
    tf = textbox(s, 1.0, top + 4.12, 11.3, 0.95)
    para(tf, "工程方法：每接入一个新模块，先写单独的测试程序确认硬件可用，再并入主固件。",
         15, True, C_ACCENT, 6, first=True)
    para(tf, "这样出问题时能立刻定位是「新模块坏了」还是「集成引入的 bug」。", 14,
         False, C_GRAY, 0)
    notes(s, "固件我们写了 11 个，不是一次写成的。每个新模块先单独验证再集成——"
             "这是我们的工程习惯，也确实帮我们快速定位过几次问题。")

    # ===== 10 服务端与仪表板（C） =====
    s = blank(prs)
    n += 1
    top = header(s, 3, "服务端与仪表板：现场报警实际效果")
    add_picture_fit(s, crop_bottom(os.path.join(ROOT, "evidence",
                                                "dashboard_live_alarm.png")),
                    0.75, top + 0.08, 6.2, 3.9)
    tf = textbox(s, 7.25, top + 0.1, 5.2, 4.2)
    para(tf, "仪表板实现的功能", 16, True, C_DARK, 8, first=True)
    para(tf, "实时：振动 RMS、被测体表面温度", 14, False, C_GRAY, 4)
    para(tf, "三级状态：正常 / 注意 / 报警", 14, False, C_GRAY, 4)
    para(tf, "维修建议：按状态给出处理建议", 14, False, C_GRAY, 4)
    para(tf, "趋势：最近 1 小时 RMS 曲线", 14, False, C_GRAY, 4)
    para(tf, "历史：SQLite 查询与曲线", 14, False, C_GRAY, 12)
    para(tf, "截图说明", 15, True, C_DARK, 6)
    para(tf, "左图为加 0.35 g 配重后的真实报警状态："
             "仪表盘变红、等级显示「报警」、建议「停机检查动平衡与轴承」，"
             "右侧趋势图记录了 RMS 由 0.06 陡升至 0.6 的阶跃。", 13, False, C_GRAY, 0)
    notes(s, "这张截图是真实报警时的界面。趋势图上那个台阶就是加配重的瞬间，"
             "这是最有说服力的一张图。")

    # ===== 11 进度对照（C） =====
    s = blank(prs)
    n += 1
    top = header(s, 3, "进度对照：主体功能已提前完成")
    rows = [["阶段", "原计划", "实际状态", "说明"],
            ["需求与方案", "第 2–3 周", "已完成", "提案 17 页、实施方案 25 页"],
            ["硬件搭建", "第 4–5 周", "已完成", "三模块全部接入并验证"],
            ["固件开发", "第 6–8 周", "已完成", "11 个固件，含 FFT、判定、存储"],
            ["系统集成", "第 9–10 周", "已完成", "MQTT 全链路 + 仪表板 + 数据库"],
            ["测试与实验", "第 11 周", "已完成主要部分", "不平衡三级、堵塞三级、温度标定"],
            ["报告与答辩", "第 12 周", "进行中", "报告大纲已完成，正文撰写中"]]
    table(s, rows, 0.8, top + 0.15, 11.7, [0.18, 0.16, 0.20, 0.46], size=13,
          row_h=0.55, header_size=13)
    rect(s, 0.8, top + 4.05, 11.7, 1.1, RGBColor(0xE8, 0xF5, 0xE9))
    tf = textbox(s, 1.0, top + 4.15, 11.3, 0.9)
    para(tf, "整体进度约提前 2 周：系统已可完整运行并演示，剩余工作集中在报告与细节打磨。",
         15, True, C_GREEN, 0, first=True)
    notes(s, "进度方面可以说得自信一点：六项里五项已完成，报告在写。"
             "系统现在就能现场演示。")

    # ===== 12 传感器标定（D） =====
    s = blank(prs)
    n += 1
    top = header(s, 4, "测试一：两路传感器的标定与验证")
    rows = [["通道", "项目", "方法", "结果", "判据"],
            ["振动", "器件真伪", "读 WHO_AM_I 原始寄存器", "0x68（原厂）", "= 0x68"],
            ["振动", "重力模长", "静止读取三轴合成", "0.991 ~ 1.002 g", "1.000 ± 0.2 g"],
            ["振动", "噪声底", "1024 点 × 多帧，带限 RMS", "0.0053 g", "—"],
            ["振动", "采样率", "实测帧时长反算", "999.7 Hz（−0.03%）", "误差 ≤ 0.5%"],
            ["温度", "稳定性", "静置 90 s", "噪声 < 1 LSB", "—"],
            ["温度", "响应性", "手指捏住探头 30 s", "+3.25 ℃，τ ≈ 25–30 s", "明显响应"],
            ["温度", "恢复性", "松手 100 s", "−4.50 ℃，无迟滞", "可逆"]]
    table(s, rows, 0.8, top + 0.12, 11.7, [0.10, 0.16, 0.30, 0.26, 0.18], size=12,
          row_h=0.46, header_size=12)
    tf = textbox(s, 0.8, top + 3.95, 11.7, 1.0)
    para(tf, "两条实测结论", 15, True, C_DARK, 6, first=True)
    para(tf, "① 采样率误差 −0.03%，比判据要求好 16 倍 —— 表明 micros() 定时与 I²C 时序可靠",
         14, False, C_GRAY, 4)
    para(tf, "② 温度时间常数 τ ≈ 30 s，意味着温度报警必须设分钟级确认延时，"
             "否则手碰一下就会误报", 14, False, C_GRAY, 0)
    notes(s, "这是第一个测试：传感器准不准。振动有四项，温度有三项，都给了具体数字和判据。")

    # ===== 13 不平衡实验（D） =====
    s = blank(prs)
    n += 1
    top = header(s, 4, "测试二：转子不平衡工况（三级）")
    add_picture_fit(s, os.path.join(ROOT, "experiments", "data", "spectra_compare.png"),
                    0.6, top + 0.05, 7.3, 3.4)
    rows = [["配重", "带内 RMS", "1× 峰值"],
            ["0 g", "0.0574 g", "0.0292 g"],
            ["0.35 g", "1.4723 g", "1.6605 g"],
            ["0.70 g", "2.6816 g", "2.8015 g"]]
    table(s, rows, 8.1, top + 0.12, 4.4, [0.30, 0.38, 0.32], size=12, row_h=0.42,
          header_size=12)
    tf = textbox(s, 8.1, top + 1.95, 4.4, 2.3)
    para(tf, "关键结论", 14, True, C_DARK, 5, first=True)
    para(tf, "0.35 g（转子质量 0.4%）使 RMS 上升 25.7 倍", 12, False, C_GRAY, 4)
    para(tf, "谐波系列（1×/3×/4×/7×）同步抬升，是不平衡的典型谱特征", 12, False, C_GRAY, 4)
    para(tf, "由 232 Hz 叶片通过频率反推叶片数为 7，与实物一致", 12, False, C_GRAY, 0)
    tf = textbox(s, 0.6, top + 3.55, 7.3, 0.7)
    para(tf, "左：三工况频谱对比（蓝=正常，橙=0.35 g，绿=0.70 g）", 12, False, C_GRAY,
         0, first=True)
    notes(s, "不平衡实验：加蓝丁胶配重，从 0 到 0.35 到 0.70 克。"
             "图上能看到 1× 处峰值一路抬高，而且谐波同步抬升——这是不平衡的指纹。")

    # ===== 14 风道堵塞实验（D） =====
    s = blank(prs)
    n += 1
    top = header(s, 4, "测试三：风道堵塞工况 —— 双通道的价值")
    rows = [["工况", "温度", "振动 RMS", "1× 峰值", "转频"],
            ["正常", "24.05 ℃", "0.0867 g", "0.0773 g", "34.5 Hz"],
            ["遮挡 50%", "24.67 ℃（+0.62）", "0.0589 g（−32%）", "0.0487 g", "34.2 Hz"],
            ["遮挡 80%", "24.85 ℃（+0.80）", "0.0605 g（−30%）", "0.0242 g（−69%）", "31.5 Hz"]]
    table(s, rows, 0.8, top + 0.12, 11.7, [0.16, 0.22, 0.22, 0.22, 0.18], size=12,
          row_h=0.46, header_size=12)

    tf = textbox(s, 0.8, top + 1.95, 5.6, 2.6)
    para(tf, "单调性检验", 15, True, C_DARK, 6, first=True)
    para(tf, "温度：单调上升 ✔", 13, False, C_GREEN, 3)
    para(tf, "1× 峰值：单调下降 ✔", 13, False, C_GREEN, 3)
    para(tf, "转频：单调下降 ✔", 13, False, C_GREEN, 3)
    para(tf, "振动 RMS：非单调 ✘（50% 与 80% 几乎相同）", 13, False, C_WARN, 0)

    rect(s, 6.7, top + 1.95, 5.85, 2.55, RGBColor(0xFD, 0xF2, 0xE9))
    tf = textbox(s, 6.95, top + 2.1, 5.4, 2.3)
    para(tf, "两类故障的特征方向相反", 15, True, C_WARN, 6, first=True)
    para(tf, "转子不平衡：温度不变、振动 ↑↑（20~65 倍）", 13, False, C_GRAY, 4)
    para(tf, "风道堵塞：温度 ↑、1× ↓、转频 ↓、振动 ↓ 32%", 13, False, C_GRAY, 8)
    para(tf, "如果只有振动传感器，堵塞会让振动下降 32%，系统会误判为"
             "「设备状态改善」，故障被完全掩盖。", 13, True, C_GRAY, 0)
    notes(s, "这页是我们最想讲的一点：两类故障的特征是反的。"
             "只装振动传感器的话，堵塞会让振动变小，系统反而以为变好了。"
             "这就是为什么要两路传感器。")

    # ===== 15 关键发现与局限（D） =====
    s = blank(prs)
    n += 1
    top = header(s, 4, "关键发现与局限：一次误判带来的经验")
    tf = textbox(s, 0.8, top + 0.1, 5.6, 4.5)
    para(tf, "发现：支撑刚度的影响达 3.7 倍", 16, True, C_DARK, 8, first=True)
    para(tf, "同一台风扇、同一转速，振动 RMS 出现 0.055 g 与 0.207 g 两种水平", 13,
         False, C_GRAY, 5)
    para(tf, "第一次判断错误：按压传感器使 RMS 下降 57%，"
             "误判为「传感器胶松了」；重新粘贴后反而升到 0.207 g", 13, False, C_GRAY, 5)
    para(tf, "原因：按压传感器时同时给整个结构加了阻尼，两个变量混淆", 13, True,
         C_WARN, 5)
    para(tf, "第二次判断正确：改为按压风扇底座，RMS 降到 0.0556 g，"
             "与历史稳定值一致 → 是风扇整体在跳动", 13, False, C_GRAY, 8)
    para(tf, "判别指标：1× 峰值 / 带内 RMS", 14, True, C_DARK, 4)
    para(tf, "固定良好 ≈ 1.0；风扇晃动 = 0.36", 13, False, C_GRAY, 0)

    rect(s, 6.9, top + 0.1, 5.65, 2.15, C_LIGHT)
    tf = textbox(s, 7.15, top + 0.25, 5.2, 1.9)
    para(tf, "当前局限（如实记录）", 15, True, C_DARK, 6, first=True)
    para(tf, "单一测点、单轴分析（三轴取最强）", 13, False, C_GRAY, 4)
    para(tf, "无转速传感器，转频靠频谱推断", 13, False, C_GRAY, 4)
    para(tf, "演示用小型风机对配重过于敏感，与工业机组可比性有限", 13, False, C_GRAY, 0)

    rect(s, 6.9, top + 2.45, 5.65, 2.1, RGBColor(0xE8, 0xF1, 0xFA))
    tf = textbox(s, 7.15, top + 2.58, 5.2, 1.85)
    para(tf, "这个发现的意义", 15, True, C_ACCENT, 6, first=True)
    para(tf, "传感器无法区分「设备自身结构振动」与「设备整体刚体运动」，"
             "工业标准要求振动测量必须相对于刚性基础，原因即在此。", 13, False, C_GRAY, 0)
    notes(s, "这页讲我们踩过的一个坑。我们先怀疑是传感器胶松了，重新贴没用；"
             "后来改成按风扇底座才找到真因。教训是：按压传感器会同时改变两个变量。")

    # ===== 16 下一步计划（D） =====
    s = blank(prs)
    n += 1
    top = header(s, 4, "下一步计划")
    rows = [["时间", "任务", "交付物"],
            ["本周", "完成报告初稿（10 章，约 30 页）", "报告初稿"],
            ["下周上半周", "补台账测试记录、整理图表", "受控文档台账"],
            ["下周下半周", "报告定稿、答辩 PPT 与演示彩排", "终稿 + 演示脚本"],
            ["答辩前", "现场演示彩排：正常 → 加配重 → 报警 → 历史查询", "可运行样机"],
            ["后续（若继续）", "复合判据写入固件、多测点组网、误报率统计", "—"]]
    table(s, rows, 0.8, top + 0.15, 11.7, [0.16, 0.56, 0.28], size=13, row_h=0.52,
          header_size=13)

    rect(s, 0.8, top + 3.55, 11.7, 1.3, RGBColor(0xE8, 0xF5, 0xE9))
    tf = textbox(s, 1.0, top + 3.66, 11.3, 1.15)
    para(tf, "现场演示方案（答辩当天）", 15, True, C_GREEN, 6, first=True)
    para(tf, "① 打开仪表板显示「正常」→ ② 现场加 0.35 g 配重 → "
             "③ 约 6 秒后仪表盘变红、显示「报警」与维修建议 → "
             "④ 切到历史页展示 SQLite 记录 → ⑤ 说明断网补发与 SD 本地存储", 13,
         False, C_GRAY, 0)
    notes(s, "最后讲计划：本周写报告，下周定稿和彩排。"
             "答辩当天我们会做现场演示，五步走，约 2 分钟。")

    # ===== 17 结束 =====
    s = blank(prs)
    n += 1
    rect(s, 0, 0, W, H, C_DARK)
    rect(s, 0, 4.55, W, 0.06, C_ACCENT)
    tf = textbox(s, 1.0, 2.35, 11.3, 1.0)
    para(tf, "汇报完毕，谢谢各位老师", 36, True, C_WHITE, 0, first=True)
    tf = textbox(s, 1.0, 3.45, 11.3, 0.6)
    para(tf, "欢迎提问与批评", 20, False, RGBColor(0xA8, 0xC4, 0xE8), 0, first=True)
    tf = textbox(s, 1.0, 4.85, 11.3, 1.4)
    para(tf, "项目仓库：github.com/Zichuan-ALT/vibration-monitor", 14, False,
         RGBColor(0xA8, 0xC4, 0xE8), 5, first=True)
    para(tf, "代码、实验记录、原始数据与证据照片均已归档，可现场查阅", 14, False,
         RGBColor(0xA8, 0xC4, 0xE8), 0)
    notes(s, "结束。如果老师问细节，仓库里都有，可以现场打开看。")

    prs.save(OUT)
    print("已生成：", OUT)
    print("幻灯片数：", len(prs.slides.__iter__.__self__._sldIdLst))
    return OUT


if __name__ == "__main__":
    build()
