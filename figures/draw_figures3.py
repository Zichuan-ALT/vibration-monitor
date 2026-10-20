# -*- coding: utf-8 -*-
"""方案 B（简化实施方案）插图：
   figS1_arch.png    简化系统架构（ESP32 单板）
   figS2_flow.png    简化固件流程
复用 draw_figures2.py 中的绘图工具与配色。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PIL import Image, ImageDraw  # noqa: E402

from draw_figures2 import (BASE, BLUE, GREEN, GREY, LBLUE, LGREEN, LGREY,  # noqa: E402
                           LORANGE, LPURPLE, LRED, ORANGE, PURPLE, RED,
                           arrow, box, elbow, label, trim)


def fig_arch():
    W, H = 1720, 1020
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    LX, LW = 30, 150
    X0, X1 = 210, W - 30

    def band(y, h, name, fill, line):
        d.rounded_rectangle([LX, y, LX + LW, y + h], radius=12, fill=fill,
                            outline=line, width=3)
        lines = name.split("\n")
        step = 28
        cy = y + h / 2 - step * (len(lines) - 1) / 2
        f = __import__("draw_figures2").font(20, True)
        for i, ln in enumerate(lines):
            tw = d.textlength(ln, font=f)
            d.text((LX + LW / 2 - tw / 2, cy + i * step - 13), ln, font=f,
                   fill=(20, 30, 45))

    band(40, 160, "应用层", LGREEN, GREEN)
    box(d, X0, 40, X1 - X0, 160,
        "Node-RED 仪表板（跑在笔记本电脑上，零成本、无需服务器）\n"
        "实时状态卡 · 历史趋势（SQLite）· 频谱 · 报警列表与维护建议 · 中英双语",
        fill=LGREEN, line=GREEN, fs=21, bold=False)

    band(260, 120, "网络层", LBLUE, BLUE)
    box(d, X0, 260, X1 - X0, 120,
        "Wi-Fi 2.4 GHz + MQTT（Broker 与仪表板装在同一台笔记本上）",
        fill=LBLUE, line=BLUE, fs=21, bold=False)

    band(440, 200, "节点层", LORANGE, ORANGE)
    box(d, X0, 440, X1 - X0, 200,
        "ESP32-WROOM-32 单板（240 MHz 双核 · 硬件 FPU · 内置 Wi-Fi）\n"
        "I²C 采振动 → 1024 点 FFT（约 30 ms）→ 特征提取\n"
        "→ 基线自学习与三级判别 → MQTT 上报（10 s 周期 + 事件即时）\n"
        "本地：microSD 断网缓存 · 板载 LED + 蜂鸣器（板载按键）",
        fill=LORANGE, line=ORANGE, fs=21, bold=False)

    band(700, 150, "感知层", LPURPLE, PURPLE)
    gap = 40
    bw = (X1 - X0 - gap) / 2
    box(d, X0, 700, bw, 150,
        "① 振动 + 辅助温度　MPU6050（GY-521）\nI²C · ±4 g · 1 kSPS · 板载温度（数字接口，无需 ADC）",
        fill=LPURPLE, line=PURPLE, fs=20, bold=False)
    box(d, X0 + bw + gap, 700, bw, 150,
        "② 表面温度　DS18B20\n1-Wire · ±0.5 ℃（数字接口，无需 ADC）",
        fill=LPURPLE, line=PURPLE, fs=20, bold=False)

    band(910, 90, "演示对象", LRED, RED)
    box(d, X0, 910, X1 - X0, 90,
        "5 V USB 小风扇固定在木板上 + 配重片（不平衡）/ 松一颗螺丝（松动）/ 挡风道（负载变化）"
        "—— 全程 5 V USB 供电，不接触市电",
        fill=LRED, line=RED, fs=20, bold=False)

    hub = (X0 + X1) / 2
    arrow(d, hub, 260, hub, 204, color=GREY)
    label(d, hub + 24, 232, "MQTT / JSON", fs=19, anchor="lm", color=GREY)
    arrow(d, hub, 440, hub, 384, color=GREY)
    for i in range(2):
        x = X0 + i * (bw + gap) + bw / 2
        elbow(d, [(x, 700), (x, 672), (hub, 672), (hub, 644)], color=PURPLE)
    img.save(os.path.join(BASE, "figS1_arch.png"))
    trim(os.path.join(BASE, "figS1_arch.png"))
    return "figS1_arch.png"


def fig_flow():
    W, H = 1560, 1360
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    cx, bw = 520, 720
    x = cx - bw / 2

    def nbox(y, text, fill=LBLUE, line=BLUE, fs=20, h=92):
        box(d, x, y, bw, h, text, fill=fill, line=line, fs=fs, bold=False)
        return y + h

    y = 60
    y = nbox(y, "上电 → 初始化（SPI / 1-Wire / Wi-Fi / SD 卡）",
             fill=LGREEN, line=GREEN)
    arrow(d, cx, y, cx, y + 34); y += 34
    y = nbox(y, "基线自学习：开机连续采集 60 s，统计振动 RMS 的均值 μ 与标准差 σ",
             fill=LORANGE, line=ORANGE, fs=19)
    arrow(d, cx, y, cx, y + 34); y += 34
    y = nbox(y, "采集一帧：振动 1024 点 + 温度 1 次", fs=19)
    arrow(d, cx, y, cx, y + 34); y += 34
    y = nbox(y, "预处理：去均值 → Hanning 窗 → 幅值恢复", fill=LGREY, line=GREY)
    arrow(d, cx, y, cx, y + 34); y += 34
    y = nbox(y, "特征提取：RMS · 峰值 · 速度有效值 · 主频与 1× 幅值 · 表面温度",
             fs=18, h=100)
    arrow(d, cx, y, cx, y + 34); y += 34
    y = nbox(y, "与基线比较：RMS > μ+3σ 为注意，> μ+6σ 为报警\n"
                "温度 > 基线 +8 ℃ 同步升级为注意，> +15 ℃ 为报警",
             fill=LRED, line=RED, fs=18, h=112)
    arrow(d, cx, y, cx, y + 34); y += 34
    y = nbox(y, "输出：三级状态（正常 / 注意 / 报警）+ 维护建议",
             fill=LGREY, line=GREY, fs=20, h=96)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "上报与记录：MQTT 发布（10 s 周期 + 事件即时）\n"
                "SD 卡缓存 · RGB LED + 蜂鸣器",
             fs=19, h=108)
    arrow(d, cx, y, cx, y + 28); y += 28
    y = nbox(y, "延时 1 s，进入下一循环（断网自动重连并补传）",
             fill=LGREEN, line=GREEN, fs=19, h=86)

    bx = 56
    elbow(d, [(x, y - 43), (bx, y - 43), (bx, 246), (x, 246)], color=GREY)
    label(d, bx + 16, 700, "循环采集", fs=20, color=GREY, bold=True, anchor="lm")
    img.save(os.path.join(BASE, "figS2_flow.png"))
    trim(os.path.join(BASE, "figS2_flow.png"))
    return "figS2_flow.png"


if __name__ == "__main__":
    print("saved:", fig_arch())
    print("saved:", fig_flow())
