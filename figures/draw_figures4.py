# -*- coding: utf-8 -*-
"""方案接线示意图（图 S2）：MPU6050 + DS18B20 + microSD 的接线方式。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PIL import Image, ImageDraw  # noqa: E402

from draw_figures2 import (BASE, BLUE, GREEN, GREY, LBLUE, LGREEN, LGREY,  # noqa: E402
                           LORANGE, LPURPLE, LRED, ORANGE, PURPLE, RED,
                           arrow, box, elbow, font, label, trim)


def fig_wiring():
    W, H = 1700, 1120
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    box(d, 100, 150, 470, 640,
        "ESP32-WROOM-32 开发板\n（带排针的成品模块）\n\n"
        "本方案用到的引脚\n"
        "GPIO21 / 22　I²C 数据 / 时钟\n"
        "GPIO26　　　MPU6050 数据就绪\n"
        "GPIO5　　　 DS18B20 数据\n"
        "GPIO18 / 19 / 23　SPI 时钟 / 数据\n"
        "GPIO13　　　microSD 片选\n"
        "GPIO25　　　蜂鸣器\n"
        "GPIO2　　　 板载 LED（无需接线）\n"
        "5V / VIN　　microSD 卡座模块电源\n"
        "GPIO0　　　 板载 BOOT 按键（无需接线）\n"
        "3V3 / GND　 公共电源",
        fill=LORANGE, line=ORANGE, fs=19, bold=False, lh=32)

    periph = [
        (150, "① MPU6050 模块（GY-521，I²C）",
         "三轴加速度 + 三轴陀螺 + 板载温度\n"
         "VCC(3V3) · GND · SDA→GPIO21 · SCL→GPIO22 · INT→GPIO26", LPURPLE, PURPLE),
        (320, "② DS18B20 温度探头（1-Wire）",
         "数据→GPIO5 · 3V3 · GND\n"
         "裸线端经螺丝端子转成排针", LPURPLE, PURPLE),
        (490, "③ microSD 卡座模块（SPI）",
         "SCK→GPIO18 · MISO→GPIO19 · MOSI→GPIO23\nCS→GPIO13 · VCC 接 5V · GND", LPURPLE, PURPLE),
        (660, "④ 有源蜂鸣器模块（3 线）",
         "VCC→3V3 · GND · 信号→GPIO25\n"
         "板载 LED 与 BOOT 按键无需任何接线", LRED, RED),
    ]
    for i, (y, name, conn, f, l) in enumerate(periph):
        box(d, 1110, y, 560, 150, f"{name}\n{conn}", fill=f, line=l, fs=18,
            bold=False, lh=28)
        yc = y + 75
        elbow(d, [(570, 470), (840, 470), (840, yc), (1106, yc)], color=GREY,
              width=3, head=11)
        label(d, 973, yc - 26, ["① 振动 + 温度（I²C）", "② 表面温度（1-Wire）",
                                "③ 存储（SPI）", "④ 声音指示"][i],
              fs=18, bold=True, color=PURPLE if i < 3 else RED)

    box(d, 100, 850, 1570, 250,
        "接线做法与供电\n"
        "所有模块均为带排针的成品板；传感器与主控之间用杜邦线母-母直插，"
        "电源用母-公线引入面包板电源轨。\n"
        "电源分配：ESP32 的 3V3 / GND / 5V 经母-公线接到面包板电源轨，"
        "各模块的 VCC / GND 从电源轨取电，避免多根线挤在同一个针脚上。\n"
        "ESP32 由 USB 充电器供电，5 V USB 风扇由另一个充电器或充电宝供电——"
        "两路分开，避免电机噪声串入数字电源。\n"
        "所有模块的 GND 与 ESP32 共地。MPU6050、DS18B20、蜂鸣器接 3V3；"
        "microSD 卡座模块 VCC 要求 4.5~5.5 V，接 ESP32 的 5V / VIN。"
        "信号线尽量短（≤ 20 cm）。",
        fill=LGREY, line=GREY, fs=18, bold=False, lh=30)
    img.save(os.path.join(BASE, "figS_wiring.png"))
    trim(os.path.join(BASE, "figS_wiring.png"))
    return "figS_wiring.png"


if __name__ == "__main__":
    print("saved:", fig_wiring())
