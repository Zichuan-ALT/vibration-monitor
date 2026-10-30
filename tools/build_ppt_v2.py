# -*- coding: utf-8 -*-
"""生成 8.1 振动状态监测器 中期答辩 PPT（纯技术版）。

与 v1 的区别：
  · 只讲技术内容——去掉汇报分工、选题背景、进度对照、演示流程、仓库等组织性页面；
  · 四人各讲一部分技术内容，每页左上角标注部分与主讲人；
  · 每页均有演讲备注。
输出：8.1_振动状态监测器_中期答辩_技术版.pptx（不覆盖已人工修改的 v1）
"""
import os

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "8.1_振动状态监测器_中期答辩_技术版.pptx")
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
    1: ("第一部分", "系统与技术指标", "成员 A", C_ACCENT),
    2: ("第二部分", "信号处理与判定", "成员 B", RGBColor(0x8E, 0x44, 0xAD)),
    3: ("第三部分", "硬件与固件实现", "成员 C", RGBColor(0x1F, 0x8A, 0x8A)),
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


def bullets(slide, left, top, width, height, items, size=15, gap=7):
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
    out = os.path.join(tempfile.gettempdir(), "ppt2_" + os.path.basename(src))
    im.crop((0, 0, w, cut)).save(out)
    return out


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, part, title, subtitle=None):
    label, part_name, speaker, color = PARTS[part]
    rect(slide, 0, 0, W, 0.09, color)
    chip = rect(slide, 0.5, 0.34, 2.75, 0.42, color)
    tf = chip.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Pt(2)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "%s · %s" % (part_name, speaker)
    set_run(r, 12.5, True, C_WHITE)

    tf = textbox(slide, 3.4, 0.26, 9.4, 0.62, MSO_ANCHOR.MIDDLE)
    para(tf, title, 24, True, C_DARK, 0, first=True)
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


def code_box(slide, lines, left, top, width, height, size=11.5):
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
    para(tf, "利莫瑞克大学 × LICHPU　工业评估（15 ECTS）· 主题 8", 13, False,
         RGBColor(0xBF, 0xD4, 0xEE), 0, first=True)
    tf = textbox(s, 0.9, 0.86, 11.6, 0.75)
    para(tf, "项目 8.1　振动状态监测器", 36, True, C_WHITE, 0, first=True)
    tf = textbox(s, 0.9, 1.6, 11.6, 0.4)
    para(tf, "中期答辩 · 技术版", 17, False, RGBColor(0xBF, 0xD4, 0xEE), 0,
         first=True)

    tf = textbox(s, 0.9, 2.7, 11.6, 2.6)
    para(tf, "本次汇报只讲技术内容，分四部分由四位成员分别主讲：", 15, False,
         C_GRAY, 12, first=True)
    for i, (label, name, spk, color) in PARTS.items():
        p = tf.add_paragraph()
        p.space_after = Pt(9)
        r = p.add_run()
        r.text = "%s　%s　——　%s" % (label, name, spk)
        set_run(r, 17, True, color)

    tf = textbox(s, 0.9, 5.55, 11.6, 0.4)
    para(tf, "小组成员：成员 A ／ 成员 B ／ 成员 C ／ 成员 D", 14, False, C_GRAY, 4,
         first=True)
    para(tf, "答辩日期：2026 年 11 月", 14, False, C_GRAY, 0)

    hline(s, 0.9, 5.42, 11.5, C_ACCENT, 1.5)
    notes(s, "开场：这次汇报不讲项目背景和分工，直接讲技术。"
             "四部分分别是系统与指标、信号处理与判定、硬件与固件、测试与数据，每人一段。")


def s_overview(prs, n):
    s = blank(prs)
    top = header(s, 1, "监测对象与系统构成")
    bullets(s, 0.55, top + 0.05, 6.35, 4.9, [
        ("被测对象：直流无刷风机 YD9225HSL", 0, True, C_DARK),
        ("DC 5 V / 0.20 A，92×92×25 mm，7 叶片，转子约 90 g", 1, False, C_GRAY),
        ("实测转速 2050 ~ 2120 RPM → 转频 34 ~ 35 Hz（铭牌标 1700 RPM）", 1, False, C_GRAY),
        ("传感与执行", 0, True, C_DARK),
        ("MPU6050：三轴加速度，I²C，0x68，±4 g，DLPF 256 Hz", 1, False, C_GRAY),
        ("DS18B20：被测体表面温度，1-Wire，12 位（0.0625 ℃）", 1, False, C_GRAY),
        ("microSD：本地记录，SPI", 1, False, C_GRAY),
        ("三层结构", 0, True, C_DARK),
        ("传感层 → 边缘层（ESP32）→ 服务层（Node-RED + SQLite）", 1, False, C_GRAY),
        ("判定在边缘完成：只上传特征值，约 180 字节/条，不上传原始波形", 1, True, C_ACCENT),
    ], size=14.5, gap=5)
    add_picture_fit(s, os.path.join(EVD, "setup_overview.jpg"), 7.15, top + 0.1,
                    5.6, 4.7)
    tf = textbox(s, 7.15, 6.15, 5.6, 0.3)
    para(tf, "实测装置：风机 + MPU6050 + DS18B20 + ESP32", 11, False,
         RGBColor(0x66, 0x66, 0x66), 0, first=True, align=PP_ALIGN.CENTER)
    pageno(s, n)
    notes(s, "先交代被测对象：一台 92 mm 的 5 V 风机，实测转速比铭牌高 21%，"
             "这对后面的基频带选择有影响。传感器两路，加一张 SD 卡。"
             "架构上最要紧的一点：判定在设备端做，只上传特征值。")


def s_metrics(prs, n):
    s = blank(prs)
    top = header(s, 1, "技术指标：设计判据与实测值")
    table(s, [
        ["指标", "设计值 / 判据", "实测结果", "结论"],
        ["采样率", "1 kSPS，误差 ≤ 0.5 %", "999.7 Hz，误差 −0.03 %", "优于判据 16 倍"],
        ["频率分辨率", "Δf = fs / N ≤ 1 Hz", "0.977 Hz（N = 1024）", "达成"],
        ["量程与灵敏度", "±4 g", "8192 LSB/g", "0.7 g 配重时峰值 3.2 g 不削顶"],
        ["带内噪声底", "—", "0.0053 g（10~260 Hz）", "与手册推算同量级"],
        ["重力模长", "1.000 ± 0.2 g", "0.991 ~ 1.002 g", "偏差 ≤ 0.9 %"],
        ["报警响应时间", "≤ 10 s", "6.0 s（3 帧 × 2 s）", "达成"],
        ["上报周期 / 报文", "≤ 5 s / ≤ 512 B", "2 s / 约 180 B", "达成"],
        ["断网数据完整性", "零丢失", "5/5 补发，时间戳偏差 0 s", "达成"],
    ], 0.55, top + 0.06, 12.2, [0.16, 0.26, 0.30, 0.28], size=12.5, row_h=0.46)
    tf = textbox(s, 0.55, 5.58, 12.2, 0.45)
    para(tf, "说明：精度与时延类指标一律取多次测量的最大值（最差情况），不取平均值。",
         12.5, False, C_ACCENT, 0, first=True)
    pageno(s, n)
    notes(s, "这张表是全部技术指标的汇总，后面各部分会逐项讲怎么得到的。"
             "请注意最后一行：断网 5 条全部补发，时间戳仍在原测量时刻，这是设计的结果不是巧合。")


def s_sampling(prs, n):
    s = blank(prs)
    top = header(s, 2, "采样参数是怎么定下来的")
    bullets(s, 0.55, top + 0.05, 6.3, 4.8, [
        ("① 采样率 fs = 1 kHz", 0, True, C_DARK),
        ("分析带宽取 10 ~ 260 Hz（低于 10 Hz 是整体运动，高于 260 Hz 超出小型风机激励范围）", 1, False, C_GRAY),
        ("Nyquist 要求 ≥ 520 Hz；取 1 kHz 留约 4 倍余量，并让 DLPF 256 Hz 的过渡带落在带宽外", 1, False, C_GRAY),
        ("② 帧长 N = 1024", 0, True, C_DARK),
        ("Δf = fs / N = 0.977 Hz；帧长 1.024 s", 1, False, C_GRAY),
        ("再长 → 分辨率更高但报警迟钝；再短 → 1× 与相邻谐波分不开", 1, False, C_GRAY),
        ("实测转频 34 Hz，相邻谐波间隔也是 34 Hz，0.977 Hz 足够区分", 1, False, C_GRAY),
        ("③ 量程 ±4 g", 0, True, C_DARK),
        ("0.7 g 配重实测峰值 3.2 g：±2 g 会削顶，±8 g 分辨率减半", 1, False, C_GRAY),
    ], size=14, gap=5)

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
    ], 7.05, top + 0.15, 5.75, 3.0, size=11)

    tf = textbox(s, 7.05, top + 3.35, 5.75, 1.5)
    para(tf, "为什么不用 delay(1)", 14, True, C_DARK, 4, first=True)
    para(tf, "delay() 每次都会叠加调度误差，累积 1024 次后可达数十毫秒；"
             "用 micros() 对齐绝对时刻，只用「等到该采的时刻」，累计误差不增长。",
         12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "这三个参数不是拍脑袋定的。采样率按 4 倍 Nyquist，帧长在分辨率与实时性之间折中，"
             "量程按实测峰值留余量。定时用 micros() 绝对对齐，这是 −0.03% 精度的来源。")


def s_rms(prs, n):
    s = blank(prs)
    top = header(s, 2, "为什么必须带限计算 RMS", "这是本项目在实现中修正的一个设计")
    bullets(s, 0.55, top + 0.05, 6.35, 4.8, [
        ("现象", 0, True, C_WARN),
        ("模块静止时，时域 RMS 仍有 0.05 g", 1, False, C_GRAY),
        ("按传感器噪声密度推算应只有 0.006 g，相差约 10 倍", 1, False, C_GRAY),
        ("原因", 0, True, C_WARN),
        ("设备并非绝对静止——桌面因人员走动、建筑振动而轻微运动", 1, False, C_GRAY),
        ("这些 < 10 Hz 的低频成分与「设备自身振动」完全是两回事，却被算进了 RMS", 1, False, C_GRAY),
        ("修正", 0, True, C_GREEN),
        ("先 FFT，再在幅度谱上按 Parseval 定理只累加 10 ~ 260 Hz 的能量", 1, False, C_GRAY),
        ("静止读数 0.0053 g（与手册一致）；运转读数 0.0574 g", 1, True, C_GREEN),
        ("结论", 0, True, C_DARK),
        ("「看起来最朴素的 RMS」也必须先明确要统计什么频段", 1, True, C_DARK),
    ], size=14, gap=4)

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
    ], 7.1, top + 0.25, 5.7, 3.2, size=11.5)

    tf = textbox(s, 7.1, top + 3.65, 5.7, 1.3)
    para(tf, "对判定的影响", 14, True, C_DARK, 4, first=True)
    para(tf, "若用全带 RMS，静止与运转只差约 10 倍，阈值与噪声底混在一起；"
             "带限后两者相差约 11 倍但噪声底更可信，正常/注意/报警三档才分得开。",
         12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "这一页是整个信号处理里最关键的一处修正。最初直接用整段时域求 RMS，"
             "静止时居然有 0.05 g，跟手册差 10 倍。查下来是桌面在动。改成频域带限积分就对了。")


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
        ("自学习时若风扇恰在停转，学到的 μ 仅 0.0053 g，推算阈值约 0.008 g", 1, False, C_GRAY),
        ("→ 设备一启动就误报。故规定：μ < 0.010 g 时拒绝该基线，状态置 uncalibrated", 1, True, C_GREEN),
    ], size=14, gap=4)

    table(s, [
        ["自学习结果", "μ", "σ", "注意阈值", "报警阈值"],
        ["第一次（固定较紧）", "0.0526 g", "0.0079 g", "0.0763 g", "0.1000 g"],
        ["第二次（支撑偏软）", "0.0942 g", "0.0112 g", "0.1279 g", "0.1676 g"],
    ], 7.15, top + 0.25, 5.65, [0.30, 0.17, 0.17, 0.18, 0.18], size=11.5, row_h=0.42)

    tf = textbox(s, 7.15, top + 1.85, 5.65, 2.6)
    para(tf, "自学习确实起了自适应作用", 14, True, C_DARK, 5, first=True)
    para(tf, "两次 μ 相差近 1.8 倍，阈值随之上移。若用固定阈值 0.100 g，"
             "第二场的正常运转就会被判为报警——这正是自学习的必要性。",
         12.5, False, C_GRAY, 6)
    para(tf, "判定效果", 14, True, C_DARK, 3)
    para(tf, "加 0.35 g 配重后 RMS 1.47 g（报警阈值的 14.7 倍），"
             "0.70 g 时 2.68 g（26.8 倍），三档被干净分开，迟滞也无抖动。",
         12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "阈值不是写死的。同一台风扇换个安装状态就能差 3.7 倍，固定阈值一定出问题。"
             "所以用自学习，并且加了一条保护：学到的基线太小就拒绝，防止停机时学到错误基线。")


def s_hw(prs, n):
    s = blank(prs)
    top = header(s, 3, "硬件接口与引脚分配")
    table(s, [
        ["外设", "接口", "信号线", "速率 / 分辨率", "备注"],
        ["MPU6050", "I²C 400 kHz", "SDA=21, SCL=22", "1 ms / 8192 LSB/g", "地址 0x68（AD0 悬空）"],
        ["DS18B20", "1-Wire", "DQ=5", "2 s / 0.0625 ℃", "板载 4.7 kΩ 上拉"],
        ["microSD", "SPI", "18 / 19 / 23 / 13", "每帧一行 CSV", "VCC 必须 5 V"],
        ["MQTT", "TCP 1883", "Wi-Fi", "2 s / 约 180 B", "QoS 1"],
    ], 0.55, top + 0.06, 12.2, [0.15, 0.16, 0.22, 0.23, 0.24], size=12, row_h=0.42)

    tf = textbox(s, 0.55, top + 2.05, 5.6, 2.4)
    para(tf, "接口设计要点", 14.5, True, C_DARK, 5, first=True)
    para(tf, "三组接口互不冲突：I²C 占 21/22，单总线占 5，SPI 占 18/19/23/13。",
         12.5, False, C_GRAY, 5)
    para(tf, "SD 模块的 VCC 必须接 5 V：板载稳压芯片压差不足，接 3.3 V 时模块不工作"
             "（这是实测踩到的坑）。", 12.5, False, C_GRAY, 0)
    add_picture_fit(s, os.path.join(FIG, "figS_wiring.png"), 6.3, top + 1.95,
                    6.45, 2.9)
    pageno(s, n)
    notes(s, "硬件这一页只讲接口。三组总线互不冲突，接线图上标得很清楚。"
             "要特别提醒一句：SD 模块必须接 5 V，接 3.3 V 是不工作的。")


def s_firmware(prs, n):
    s = blank(prs)
    top = header(s, 3, "固件结构与三个可靠性机制")
    bullets(s, 0.55, top + 0.05, 6.3, 2.2, [
        ("主循环结构：先测量，后联网", 0, True, C_DARK),
        ("每轮先完成采样 → FFT → 特征 → 判定，把结果存入环形缓冲，再处理网络", 1, False, C_GRAY),
        ("这样即使 MQTT 重连阻塞数秒，也不会漏掉本周期的采样点", 1, False, C_GRAY),
    ], size=14, gap=4)

    table(s, [
        ["机制", "实现", "解决的问题"],
        ["环形缓冲 + 补发", "60 条环形缓冲（约 2 min），恢复后按序补发",
         "网络中断期间数据不丢"],
        ["NTP 时间戳", "报文 ts 取测量时刻，不取发送时刻",
         "补发数据仍落在正确的时间位置"],
        ["SD 本地记录", "每帧追加一行并 flush；报警瞬间另存 1024 点波形",
         "掉电最多丢 1 条；报警可事后做频谱诊断"],
        ["基线下限保护", "自学习 μ < 0.010 g 时拒绝基线，置 uncalibrated",
         "停机时学到错误基线导致误报"],
    ], 0.55, top + 2.35, 12.2, [0.17, 0.46, 0.37], size=11.5, row_h=0.5)

    tf = textbox(s, 0.55, top + 4.75, 12.2, 0.5)
    para(tf, "四个机制互相独立：任一环节失效时系统降级而不断链——"
             "例如 SD 卡故障只关闭本地记录，联网上报不受影响。", 12.5, True, C_ACCENT, 0,
         first=True)
    pageno(s, n)
    notes(s, "固件这边讲三件事：主循环顺序、四个可靠性机制、以及降级设计。"
             "核心思想是「先测量后联网」，网络问题不能影响采样。")


def s_issues(prs, n):
    s = blank(prs)
    top = header(s, 3, "实现中解决的三个技术问题")
    items = [
        ("① I²C 在 1 kHz 采样下偶发丢帧", C_ACCENT,
         ["现象：连续采样 100 帧，偶有无效帧，数据不连续",
          "定位：1 ms 周期内读 14 字节（含温度）超出 I²C 时间开销",
          "处理：I²C 提至 400 kHz，且只读 6 字节加速度（温度另开 2 s 周期）",
          "结果：丢帧率降到 0"]),
        ("② arduinoFFT 2.0.4 API 与 1.x 教程不同", C_WARN,
         ["现象：照抄 1.x 写法直接编译失败",
          "原因：2.x 改为模板类 ArduinoFFT<T>，方法名全小写",
          "处理：ArduinoFFT<double> FFT(...)；windowing(FFTWindow::Hann, ...)",
          "结果：编译通过，频谱与预期一致"]),
        ("③ 库对 WHO_AM_I 做了归一化", C_GREEN,
         ["现象：getDeviceID() 返回 0x34 而非手册的 0x68，疑似假芯片",
          "原因：该库取 >>1 后的低 6 位，寄存器原值确为 0x68",
          "处理：自检程序同时读 0x75 原始值与库返回值，两者都核对",
          "结果：确认为原厂器件"]),
    ]
    y = top + 0.1
    for title, color, lines in items:
        rect(s, 0.55, y, 0.08, 1.55, color)
        tf = textbox(s, 0.8, y, 12.0, 1.55)
        para(tf, title, 15.5, True, color, 4, first=True)
        for ln in lines:
            para(tf, ln, 12.5, False, C_GRAY, 2)
        y += 1.72
    pageno(s, n)
    notes(s, "这三个问题都是实现中真实遇到的，也都解决了。第二个和第三个属于「文档与实现不一致」，"
             "经验是：关键判据要直读底层寄存器，第三方封装不能全信。")


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
             "速率单调递减 → 典型一阶响应，τ ≈ 25~30 s。"
             "结论：温度报警必须设分钟级确认延时，不能照搬振动通道的秒级迟滞，"
             "否则手碰一下探头就会误报。", 12.5, False, C_GRAY, 0)
    pageno(s, n)
    notes(s, "标定分两个通道。振动做了四项自洽性检查，温度做了稳定性/响应/恢复三项。"
             "重点是温度的时间常数约 30 秒，这直接决定了温度报警不能用秒级迟滞。")


def s_imb(prs, n):
    s = blank(prs)
    top = header(s, 4, "故障实验一：转子不平衡", "单片扇叶贴蓝丁胶配重，距轴心约 35 mm")
    table(s, [
        ["工况", "配重", "带内 RMS", "1× 峰值", "峰值频率"],
        ["① 基准", "0 g", "0.0574 g", "0.0292 g", "32.2 Hz"],
        ["② 轻微", "0.35 g", "1.4723 g（+25.7 倍）", "1.6605 g（+56.9 倍）", "33.2 Hz"],
        ["③ 明显", "0.70 g", "2.6816 g（+46.7 倍）", "2.8015 g（+95.9 倍）", "31.2 Hz"],
    ], 0.55, top + 0.04, 6.35, [0.16, 0.16, 0.27, 0.26, 0.15], size=11, row_h=0.4)

    tf = textbox(s, 0.55, top + 1.7, 6.35, 2.9)
    para(tf, "灵敏度", 14, True, C_DARK, 3, first=True)
    para(tf, "0.35 g 仅占转子质量（约 90 g）的 0.4 %，却使 RMS 上升 25.7 倍、"
             "1× 峰值上升 56.9 倍。基准工况的带内 RMS 已是噪声底的约 10 倍，"
             "加配重后拉大到约 270 倍，轻微劣化也在可检测范围内。",
         12.5, False, C_GRAY, 6)
    para(tf, "亚线性与共振", 14, True, C_DARK, 3)
    para(tf, "配重加倍只带来 1.82 倍响应。用 F = m·r·ω² 校核：0.35 g 在 35 mm、32 Hz 下"
             "约 0.49 N，对应刚体加速度 0.56 g，而实测峰值 1.66 g，放大倍数 Q ≈ 3.7 —— "
             "系统工作在共振区附近，响应被阻尼压制。", 12.5, False, C_GRAY, 0)

    add_picture_fit(s, os.path.join(DAT, "spectra_compare.png"), 7.0, top + 0.1,
                    5.75, 4.9)
    pageno(s, n)
    notes(s, "不平衡实验是本项目的核心验证。三级工况，每档 6 到 7 帧。"
             "右边频谱图三条曲线在 33 Hz 附近依次抬升，谐波同步上升——这是不平衡的典型谱特征。")


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
    para(tf, "风扇振动包含机械与气动两部分。堵塞使风量减少、气动激励减弱，"
             "总振动随之下降；同时电机负载增大导致转速下降（34.5 → 31.5 Hz），"
             "散热变差导致温度上升。", 12.5, False, C_GRAY, 7)
    para(tf, "单以 RMS 作判据不可靠", 14.5, True, C_WARN, 4)
    para(tf, "RMS 在 50 % 与 80 % 之间几乎不变（0.0589 与 0.0605），不随堵塞程度单调变化。"
             "真正单调的是温度、1× 峰值与转频三项。", 12.5, False, C_GRAY, 0)

    add_picture_fit(s, os.path.join(EVD, "blockage_test.jpg"), 6.6, top + 1.85,
                    6.2, 3.4)
    pageno(s, n)
    notes(s, "堵塞实验的结论和直觉相反：堵塞越严重，振动越小。因为气动激励减弱了。"
             "所以单看 RMS 会把这种劣化误判成「状态改善」。真正单调的是温度、1× 峰值、转频。")


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
        ("重新粘贴后反而升到 0.207 g —— 按压时手指同时给整个结构加了阻尼，两个变量混淆", 1, False, C_GRAY),
        ("第二次判断（正确）", 0, True, C_GREEN),
        ("改为按压风扇底座，手指远离传感器 → RMS 降到 0.0556 g（−73 %）", 1, False, C_GRAY),
        ("结论：不是传感器松了，而是风扇整体在桌面上跳动，刚体运动被一并测到", 1, True, C_GREEN),
    ], size=13, gap=3)

    tf = textbox(s, 6.85, top + 1.65, 5.95, 3.4)
    para(tf, "可用的在线判别指标", 14.5, True, C_DARK, 5, first=True)
    para(tf, "把「1× 峰值 / 带内 RMS」作为辅助诊断量：比值 ≈ 1.0 表示能量集中在转频、测量可信；"
             "比值显著小于 1 说明叠加了刚体运动，当前数据不可用于判故障。",
         12.5, False, C_GRAY, 8)
    para(tf, "工程意义", 14.5, True, C_DARK, 4)
    para(tf, "支撑刚度对振动测量的影响（3.7 倍）远大于传感器粘接质量。"
             "根本原因是传感器无法区分「设备自身的结构振动」与「设备整体的刚体运动」——"
             "这也是工业标准要求振动测量必须相对于刚性基础进行的原因。",
         12.5, False, C_GRAY, 8)
    para(tf, "方法论：怀疑某个因素时，应设计只改变该因素的实验；"
             "若两个变量同时被改变，即便结果符合预期也不能作为证据。",
         12.5, True, C_ACCENT, 0)
    pageno(s, n)
    notes(s, "这一页是这次项目最有价值的发现，也是一次误判的记录。"
             "第一次按压传感器，结论指向「胶松了」，但重贴后反而更差。"
             "第二次改为按压底座才定位到真正原因：风扇在跳。支撑刚度影响 3.7 倍。")


def s_limit(prs, n):
    s = blank(prs)
    top = header(s, 1, "技术局限与下一步工作")
    tf = textbox(s, 0.55, top + 0.05, 6.0, 4.8)
    para(tf, "当前技术局限", 15, True, C_WARN, 7, first=True)
    for t in [
        "单一测点、单轴分析（三轴取能量最强轴），无法做相位分析或故障定位",
        "无转速传感器，转频靠频谱推断；Δf = 0.977 Hz 限制了转速变化的观测精度",
        "温度未做绝对标定（缺 0 ℃ 一级参考点），精度依赖出厂 ±0.5 ℃",
        "电流通道未实现，题目要求的三类量只做了两类",
        "未做长时间误报率统计，目前仅累计数小时运行记录",
        "温度尚未纳入判定，只做展示与记录",
    ]:
        para(tf, "· " + t, 12.5, False, C_GRAY, 5)

    tf = textbox(s, 6.85, top + 0.05, 5.95, 4.8)
    para(tf, "下一步（按优先级）", 15, True, C_GREEN, 7, first=True)
    steps = [
        ("复合判据写入固件", "实现「温度升 + 振动降 → 气动故障」与「振动显著上升 → 机械故障」"
         "的分支判断；温度单独设分钟级延时"),
        ("长时间运行与误报率", "连续运行 24 h 以上，统计三态分布与误报次数"),
        ("温度绝对标定", "用冰水混合物（0 ℃）作一级参考点，必要时做线性修正"),
        ("转速测量", "增加转速传感器或加长帧提高分辨率，以便做阶次分析"),
        ("补充故障类型", "轴承损伤、松动、不对中的特征数据，尝试分类"),
    ]
    for i, (name, desc) in enumerate(steps, 1):
        p = tf.add_paragraph()
        p.space_after = Pt(7)
        r = p.add_run()
        r.text = "%d. %s" % (i, name)
        set_run(r, 13.5, True, C_DARK)
        para(tf, "　　" + desc, 12.5, False, C_GRAY, 4)
    pageno(s, n)
    notes(s, "最后如实交代局限。最要紧的两条：一是电流通道没做，题目要求的三类量只做了两类；"
             "二是温度还没进判定，只做展示。下一步第一优先级就是把温度做成复合判据。")


def s_end(prs):
    s = blank(prs)
    rect(s, 0, 0, W, 7.5, C_DARK)
    tf = textbox(s, 1.2, 2.3, 11.0, 1.0)
    para(tf, "技术结论", 15, False, RGBColor(0xBF, 0xD4, 0xEE), 10, first=True)
    para(tf, "双通道 + 边缘判定，可以把两类方向相反的故障分开", 30, True, C_WHITE, 8)
    para(tf, "关键是：带限 RMS、自学习阈值、以及时间戳一致的补发机制", 17, False,
         RGBColor(0xBF, 0xD4, 0xEE), 0)
    hline(s, 1.2, 4.35, 6.5, RGBColor(0xBF, 0xD4, 0xEE), 1.2)
    tf = textbox(s, 1.2, 4.6, 11.0, 1.2)
    para(tf, "谢谢，欢迎提问", 24, True, C_WHITE, 8, first=True)
    para(tf, "小组成员：成员 A ／ 成员 B ／ 成员 C ／ 成员 D", 14, False,
         RGBColor(0xBF, 0xD4, 0xEE), 0)
    notes(s, "收尾一句话：这套系统的技术价值在于把两类方向相反的故障分开，"
             "靠的是带限 RMS、自学习阈值和时间戳一致的补发。")


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    s_cover(prs)
    s_overview(prs, 2)
    s_metrics(prs, 3)
    s_sampling(prs, 4)
    s_rms(prs, 5)
    s_decision(prs, 6)
    s_hw(prs, 7)
    s_firmware(prs, 8)
    s_issues(prs, 9)
    s_calib(prs, 10)
    s_imb(prs, 11)
    s_blk(prs, 12)
    s_stiff(prs, 13)
    s_limit(prs, 14)
    s_end(prs)

    prs.save(OUT)
    print("已生成：", OUT)
    print("页数：", len(prs.slides))


if __name__ == "__main__":
    build()
