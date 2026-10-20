# -*- coding: utf-8 -*-
"""生成《震动状态监测器设计方案》插图：
   fig1_system.png    系统总体框图
   fig2_signal.png    信号调理与采集链路
   fig3_software.png  软件主流程
"""
import math
import os
from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = r"C:\Windows\Fonts\msyh.ttc"
FONT_BOLD = r"C:\Windows\Fonts\msyhbd.ttc"

BLUE, LBLUE = (31, 78, 121), (222, 235, 247)
ORANGE, LORANGE = (196, 106, 20), (252, 235, 214)
GREEN, LGREEN = (46, 116, 70), (223, 240, 226)
GREY, LGREY = (90, 96, 104), (238, 240, 243)
RED, LRED = (170, 48, 48), (250, 226, 226)


def font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_PATH, size)


def box(d, x, y, w, h, text, *, fill=LBLUE, line=BLUE, fs=22, bold=True,
        radius=14, width=3, tcolor=(20, 30, 45), lh=None):
    d.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=fill,
                        outline=line, width=width)
    f = font(fs, bold)
    lines = text.split("\n")
    step = lh if lh else fs + 8
    total = step * (len(lines) - 1)
    cy = y + h / 2 - total / 2
    for i, ln in enumerate(lines):
        tw = d.textlength(ln, font=f)
        d.text((x + w / 2 - tw / 2, cy + i * step - fs * 0.62), ln,
               font=f, fill=tcolor)


def arrow(d, x1, y1, x2, y2, *, color=GREY, width=4, head=13):
    d.line([x1, y1, x2, y2], fill=color, width=width)
    ang = math.atan2(y2 - y1, x2 - x1)
    for s in (-1, 1):
        a = ang + s * 2.60
        d.line([x2, y2, x2 + head * math.cos(a), y2 + head * math.sin(a)],
               fill=color, width=width)


def elbow(d, pts, **kw):
    for i in range(len(pts) - 1):
        if i == len(pts) - 2:
            arrow(d, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], **kw)
        else:
            d.line([pts[i], pts[i + 1]], fill=kw.get("color", GREY),
                   width=kw.get("width", 4))


def label(d, x, y, text, *, fs=19, color=(70, 76, 84), bold=False, anchor="mm"):
    d.text((x, y), text, font=font(fs, bold), fill=color, anchor=anchor)


def title(d, w, text):
    """图题由 Word 题注统一给出，插图中不再重复绘制。"""
    return None


def trim(path, pad=22):
    """裁掉四周空白，保证插图紧凑（图题由 Word 题注提供）。"""
    im = Image.open(path).convert("RGB")
    w, h = im.size
    px = im.load()
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            if px[x, y] != (255, 255, 255):
                xs.append(x)
                ys.append(y)
    if not xs:
        return
    box = (max(0, min(xs) - pad), max(0, min(ys) - pad),
           min(w, max(xs) + pad + 1), min(h, max(ys) + pad + 1))
    im.crop(box).save(path)
    print("   trimmed", path, im.size, "->", (box[2] - box[0], box[3] - box[1]))


# ------------------------------------------------------------ 图 1 系统总体框图
def fig_system():
    W, H = 1800, 1180
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    title(d, W, "图 1  震动状态监测器系统总体框图")

    box(d, 620, 88, 560, 92,
        "被测对象：电动机 / 风机 / 水泵 / 齿轮箱 / 机床主轴", fs=22)

    box(d, 150, 236, 700, 116,
        "主通道  数字 MEMS 加速度计 ADXL345\n"
        "SPI · ±16 g · ODR 3.2 kHz · 三轴 · 3.9 mg/LSB",
        fill=LGREEN, line=GREEN, fs=21)
    box(d, 950, 236, 700, 116,
        "扩展通道  IEPE 压电加速度计\n"
        "恒流源 4 mA / 24 V + 交流耦合 + 二级放大 + 抗混叠",
        fill=LORANGE, line=ORANGE, fs=21)
    elbow(d, [(740, 180), (740, 208), (500, 208), (500, 236)])
    elbow(d, [(1060, 180), (1060, 208), (1300, 208), (1300, 236)])

    box(d, 620, 404, 560, 100,
        "信号调理与抗混叠滤波\n"
        "二阶巴特沃斯低通 fc ≈ 1.5 kHz · 直流偏置 1.65 V",
        fill=LGREY, line=GREY, fs=21)
    elbow(d, [(500, 352), (500, 380), (760, 380), (760, 404)])
    elbow(d, [(1300, 352), (1300, 380), (1040, 380), (1040, 404)])

    box(d, 520, 556, 760, 120,
        "主控 STM32F407VET6（168 MHz · FPU · DSP 指令）\n"
        "12-bit ADC + DMA 双缓冲 · SPI 突发采集 · CMSIS-DSP 2048 点 FFT",
        fill=LBLUE, line=BLUE, fs=21)
    arrow(d, 900, 504, 900, 556)

    outs = [
        ("OLED 0.96\" 显示\n波形 / 频谱 / 特征值 / 状态", LGREY, GREY),
        ("声光报警\n蜂鸣器 + 红黄绿三色 LED", LRED, RED),
        ("microSD 存储\n原始数据 + 特征值 CSV", LGREY, GREY),
        ("ESP8266 WiFi\nJSON 上报 → 上位机 / 云", LGREEN, GREEN),
    ]
    bw, gap, y0, bh = 372, 44, 800, 122
    x0 = (W - (4 * bw + 3 * gap)) / 2
    for i, (t, f, l) in enumerate(outs):
        bx = x0 + i * (bw + gap)
        box(d, bx, y0, bw, bh, t, fill=f, line=l, fs=20, bold=False)
        elbow(d, [(900, 676), (900, 744), (bx + bw / 2, 744), (bx + bw / 2, y0)])

    box(d, 560, 990, 680, 96,
        "电源模块  3.7 V / 2000 mAh 锂电 + TP4056 充电管理 + 3.3 V LDO\n"
        "+ MT3608 升压 24 V 供 IEPE 恒流源（续航 ≥ 8 h）",
        fill=LORANGE, line=ORANGE, fs=20)
    elbow(d, [(560, 1038), (300, 1038), (300, 616), (520, 616)])
    label(d, 324, 952, "3.3 V / 24 V 供电", fs=20, color=ORANGE, bold=True,
          anchor="lm")

    img.save(os.path.join(BASE, "fig1_system.png"))
    return "fig1_system.png"


# ------------------------------------------------------------ 图 2 信号链
def fig_signal():
    W, H = 1800, 720
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    title(d, W, "图 2  信号调理、采集与处理链路")

    stages = [
        ("加速度计\nADXL345\n±16 g / 3.2 kHz", LGREEN, GREEN),
        ("SPI 突发读\nFIFO 流模式\nINT1 数据就绪", LBLUE, BLUE),
        ("去均值 + 去趋势\nHanning 窗\n幅值恢复 ×2", LGREY, GREY),
        ("2048 点 FFT\nCMSIS-DSP\nΔf = 1.5625 Hz", LBLUE, BLUE),
        ("特征提取\nRMS / 峰值 / 峭度\n速度有效值 / 包络谱", LORANGE, ORANGE),
        ("状态判别\n三级阈值 + 迟滞\n正常 / 注意 / 报警", LRED, RED),
    ]
    bw, gap, y0, bh = 258, 34, 236, 190
    x0 = (W - (len(stages) * bw + (len(stages) - 1) * gap)) / 2
    for i, (t, f, l) in enumerate(stages):
        x = x0 + i * (bw + gap)
        box(d, x, y0, bw, bh, t, fill=f, line=l, fs=19, bold=False)
        if i:
            arrow(d, x - gap + 4, y0 + bh / 2, x - 4, y0 + bh / 2)

    label(d, W / 2, 520,
          "采样：fs = 3.2 kSPS，帧长 N = 2048，帧重叠 50 %，"
          "分析带宽 5 Hz ~ 1.6 kHz，单帧处理耗时 < 20 ms",
          fs=22, color=(60, 66, 74), bold=True)
    label(d, W / 2, 578,
          "换算：a [m/s²] = 9.80665 × a [g]　|　"
          "速度有效值 v_rms [mm/s] = √( Σ (A_k / 2πf_k)² / 2 )，积分区间 10 ~ 1000 Hz",
          fs=20, color=GREY)
    label(d, W / 2, 634,
          "包络解调：带通 500 ~ 1500 Hz → Hilbert 包络 → FFT → "
          "识别 BPFO / BPFI / BSF / FTF 及其谐波",
          fs=20, color=GREY)

    img.save(os.path.join(BASE, "fig2_signal.png"))
    return "fig2_signal.png"


# ------------------------------------------------------------ 图 3 软件流程
def fig_software():
    W, H = 1600, 1500
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    title(d, W, "图 3  软件主流程")

    cx, bw, bh = 500, 680, 88
    x = cx - bw / 2

    def nbox(y, text, fill=LBLUE, line=BLUE, fs=21, h=bh):
        box(d, x, y, bw, h, text, fill=fill, line=line, fs=fs, bold=False)
        return y + h

    y = 92
    y = nbox(y, "开始：上电复位", fill=LGREEN, line=GREEN)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "系统初始化：时钟 168 MHz / SPI1 / ADC+DMA / I²C1 / SDIO / "
                "USART / 定时器 / OLED", fill=LGREY, line=GREY, fs=19)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "基线自学习：空载采集 30 s，统计正常工况 RMS / 峭度基线，"
                "自动生成三级阈值", fill=LORANGE, line=ORANGE, fs=19)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "采集一帧：2048 点（TIM 触发 ADC + DMA 双缓冲 / ADXL345 FIFO 突发读）",
             fs=19)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "预处理：去均值 → 去趋势 → Hanning 加窗",
             fill=LGREY, line=GREY, fs=20)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "FFT 与特征提取：RMS、峰值、峰峰值、峰值因子、峭度、速度有效值、包络谱",
             fs=19)
    arrow(d, cx, y, cx, y + 36); y += 36

    dw, dh = 460, 200
    dcy = y + dh / 2
    d.polygon([(cx, y), (cx + dw / 2, dcy), (cx, y + dh), (cx - dw / 2, dcy)],
              fill=LRED, outline=RED, width=3)
    label(d, cx, dcy - 20, "越限？", fs=28, color=(120, 30, 30), bold=True)
    label(d, cx, dcy + 30, "（连续 3 帧越限才确认）", fs=17, color=(120, 30, 30))
    y += dh

    nx0, nw = 920, 600
    box(d, nx0, 1080, nw, 140,
        "OLED 正常显示\n特征值写入 SD 卡\nWiFi 周期上报", fill=LGREEN, line=GREEN,
        fs=21, bold=False)
    elbow(d, [(cx + dw / 2, dcy), (nx0 + nw / 2, dcy), (nx0 + nw / 2, 1080)])
    label(d, cx + dw / 2 + 60, dcy - 32, "否", fs=22, color=GREEN, bold=True)

    arrow(d, cx, y, cx, y + 40)
    label(d, cx - 48, y + 22, "是", fs=22, color=RED, bold=True)
    y += 40
    y = nbox(y, "报警动作：蜂鸣器鸣响 + 红灯点亮 + OLED 显示状态等级",
             fill=LRED, line=RED, fs=20, h=96)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "保存故障快照：原始 2048 点 + 特征值 + 时间戳，并立即 WiFi 上报",
             fill=LRED, line=RED, fs=19, h=96)
    arrow(d, cx, y, cx, y + 32); y += 32
    y = nbox(y, "延时 1 s 后进入下一循环", fill=LGREY, line=GREY, fs=20, h=80)

    bx = 60
    elbow(d, [(x, y - 40), (bx, y - 40), (bx, 496), (x, 496)], color=GREY)
    label(d, bx + 18, 660, "循环采集", fs=20, color=GREY, bold=True, anchor="lm")

    img.save(os.path.join(BASE, "fig3_software.png"))
    return "fig3_software.png"


if __name__ == "__main__":
    for fn in (fig_system, fig_signal, fig_software):
        name = fn()
        print("saved:", name)
        trim(os.path.join(BASE, name))
