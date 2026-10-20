# -*- coding: utf-8 -*-
"""生成《项目 8.1 振动状态监测器》提案与需求规格说明的插图。

fig1_architecture.png  系统总体架构（感知 / 边缘 / 网络 / 应用）
fig2_datalink.png      信号链与数据流
fig3_software.png      固件主流程
fig4_dashboard.png     仪表板线框图
fig5_gantt.png         12 周计划与关口
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
PURPLE, LPURPLE = (94, 62, 145), (233, 226, 245)


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


def trim(path, pad=22):
    im = Image.open(path).convert("RGB")
    w, h = im.size
    px = im.load()
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            if px[x, y] != (255, 255, 255):
                xs.append(x)
                ys.append(y)
    if xs:
        im.crop((max(0, min(xs) - pad), max(0, min(ys) - pad),
                 min(w, max(xs) + pad + 1),
                 min(h, max(ys) + pad + 1))).save(path)


# ---------------------------------------------------------------- 图 1 架构
def fig_arch():
    W, H = 1820, 1150
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    LX, LW = 30, 168
    X0, X1 = 224, W - 30

    def band(y, h, name, fill, line):
        d.rounded_rectangle([LX, y, LX + LW, y + h], radius=12, fill=fill,
                            outline=line, width=3)
        f = font(21, True)
        lines = name.split("\n")
        step = 29
        cy = y + h / 2 - step * (len(lines) - 1) / 2
        for i, ln in enumerate(lines):
            tw = d.textlength(ln, font=f)
            d.text((LX + LW / 2 - tw / 2, cy + i * step - 14), ln, font=f,
                   fill=(20, 30, 45))

    bw, gap = 496, 34
    # 应用层
    band(40, 180, "应用层", LGREEN, GREEN)
    for i, (t, f, l) in enumerate([
        ("实时仪表板（中英双语）\n设备状态卡 · 趋势曲线 · 频谱图 · 阈值线", LGREEN, GREEN),
        ("历史与报警\n报警列表 · 事件回放 · 维护建议 · 工单导出", LGREEN, GREEN),
        ("数据出口\nMQTT 订阅 · REST/CSV · 数字孪生接入点", LGREEN, GREEN),
    ]):
        box(d, X0 + i * (bw + gap), 40, bw, 180, t, fill=f, line=l, fs=20,
            bold=False)
    # 网络层
    band(280, 140, "网络与\n数据层", LBLUE, BLUE)
    for i, (t, f, l) in enumerate([
        ("MQTT Broker\nMosquitto（TLS 可选）", LBLUE, BLUE),
        ("时序数据库\nInfluxDB / SQLite", LBLUE, BLUE),
        ("规则与告警\nNode-RED · 邮件/看板推送", LBLUE, BLUE),
    ]):
        box(d, X0 + i * (bw + gap), 280, bw, 140, t, fill=f, line=l, fs=20,
            bold=False)
    # 边缘层
    band(490, 200, "边缘\n处理层", LORANGE, ORANGE)
    box(d, X0, 490, 1230, 200,
        "主控 STM32F407VET6（168 MHz · 硬件 FPU · DSP 指令）\n"
        "多传感同步采集（SPI / 1-Wire / 外部中断）\n"
        "振动 FFT + 包络谱 + 阶次谱 · 温度趋势 · 转速稳定性\n"
        "工况门控（转速稳态判定）→ 基线自学习 → 多指标偏离度 → 三级判别",
        fill=LORANGE, line=ORANGE, fs=20, bold=False)
    box(d, X0 + 1254, 490, X1 - (X0 + 1254), 200,
        "本地人机与缓存\nOLED 状态显示\n按键配置\n蜂鸣器 + 三色 LED\nmicroSD 断网缓存",
        fill=LGREY, line=GREY, fs=19, bold=False)
    # 感知层
    band(760, 160, "感知层", LPURPLE, PURPLE)
    sw = (X1 - X0 - 2 * gap) / 3
    senses = [
        ("① 振动\nADXL345 · SPI · ±16 g\n3.2 kSPS · 3 轴", LPURPLE, PURPLE),
        ("② 温度\nDS18B20 · 1-Wire\n轴承座 / 外壳 · ±0.5 ℃", LPURPLE, PURPLE),
        ("③ 转速\n反射式光电 / 霍尔\n每转 1 脉冲 → 转频 f_r", LPURPLE, PURPLE),
    ]
    for i, (t, f, l) in enumerate(senses):
        box(d, X0 + i * (sw + gap), 760, sw, 160, t, fill=f, line=l, fs=20,
            bold=False)
    # 演示对象
    band(970, 140, "演示\n对象", LRED, RED)
    box(d, X0, 970, X1 - X0, 140,
        "安全低压演示台架（≤ 24 V DC）：小型风机 / 直流电机 + 可装配重块（不平衡）\n"
        "+ 可调松动底座（结构松动）+ 风道遮挡板（负载变化）——全程不接触市电",
        fill=LRED, line=RED, fs=20, bold=False)

    hub = X0 + 615
    for i in range(3):
        x = X0 + i * (sw + gap) + sw / 2
        elbow(d, [(x, 760), (x, 730), (hub, 730), (hub, 694)], color=PURPLE)
    arrow(d, hub, 490, hub, 424, color=GREY)
    label(d, hub + 22, 457, "Wi-Fi 2.4 GHz · MQTT / JSON · 10 s 周期 + 事件即时",
          fs=19, anchor="lm", color=GREY)
    arrow(d, hub, 280, hub, 224, color=GREY)
    img.save(os.path.join(BASE, "fig1_architecture.png"))
    trim(os.path.join(BASE, "fig1_architecture.png"))


# ---------------------------------------------------------------- 图 2 数据链
def fig_datalink():
    W, H = 1820, 810
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    rows = [
        ("感知", [("振动 ADXL345", LGREEN, GREEN), ("温度 DS18B20", LGREEN, GREEN),
                  ("转速 光电/霍尔", LGREEN, GREEN)]),
        ("采集", [("SPI + DMA 突发读\n三轴去均值", LBLUE, BLUE),
                  ("1-Wire 轮询\n每 2 s 一次", LBLUE, BLUE),
                  ("外部中断捕获\n周期 → 转速", LBLUE, BLUE)]),
        ("特征", [("RMS · 峰值 · 峭度\n速度有效值", LORANGE, ORANGE),
                  ("趋势 · 变化率\n与环境温差", LORANGE, ORANGE),
                  ("平均转速 · 波动\n启停/稳态判定", LORANGE, ORANGE)]),
    ]
    bw, gap = 300, 26
    for r, (name, cells) in enumerate(rows):
        y = 80 + r * 190
        box(d, 30, y, 130, 150, name, fill=LGREY, line=GREY, fs=22)
        for i, (t, f, l) in enumerate(cells):
            x = 210 + i * (bw + gap)
            box(d, x, y, bw, 150, t, fill=f, line=l, fs=19, bold=False)
            if r:
                arrow(d, x + bw / 2, y - 40 + 4, x + bw / 2, y - 4)

    # 汇合到融合决策
    for i in range(3):
        x = 210 + i * (bw + gap) + bw / 2
        elbow(d, [(x, 610), (x, 628), (700, 628), (700, 646)])
    box(d, 210, 646, 980, 120,
        "特征融合与门控：仅在转速稳定、温度处于稳态区间时评价\n"
        "按转速分档建立基线 → 多指标偏离度 → 三级状态（正常 / 注意 / 报警）",
        fill=LRED, line=RED, fs=19, bold=False)
    box(d, 1250, 646, 540, 120,
        "输出\nMQTT 上报 · 本地报警\nSD 缓存 · 维护建议", fill=LGREEN,
        line=GREEN, fs=19, bold=False)
    arrow(d, 1196, 706, 1246, 706)
    img.save(os.path.join(BASE, "fig2_datalink.png"))
    trim(os.path.join(BASE, "fig2_datalink.png"))


# ---------------------------------------------------------------- 图 3 流程
def fig_software():
    W, H = 1680, 1560
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    cx, bw = 560, 760
    x = cx - bw / 2

    def nbox(y, text, fill=LBLUE, line=BLUE, fs=20, h=86):
        box(d, x, y, bw, h, text, fill=fill, line=line, fs=fs, bold=False)
        return y + h

    y = 60
    y = nbox(y, "上电复位 → 看门狗与外设初始化（时钟 / SPI / 1-Wire / 中断 / Wi-Fi）",
             fill=LGREEN, line=GREEN, fs=19)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "基线自学习：在稳定转速下采集 5 min，建立振动 / 温度 / 转速分档基线",
             fill=LORANGE, line=ORANGE, fs=19)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "采集一帧：振动 2048 点 + 温度 + 转速脉冲计数", fs=19)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "预处理：去均值 → 去趋势 → Hanning 窗 → 幅值恢复", fill=LGREY,
             line=GREY, fs=19)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "特征提取：RMS / 峰值 / 峭度 / 速度有效值 / 1× 与叶频幅值 / 包络谱",
             fs=19)
    arrow(d, cx, y, cx, y + 34); y += 34

    dw, dh = 470, 190
    dcy = y + dh / 2
    d.polygon([(cx, y), (cx + dw / 2, dcy), (cx, y + dh), (cx - dw / 2, dcy)],
              fill=LRED, outline=RED, width=3)
    label(d, cx, dcy - 18, "转速稳定？", fs=25, color=(120, 30, 30), bold=True)
    label(d, cx, dcy + 26, "（否则视为瞬态，不评价）", fs=17, color=(120, 30, 30))
    y += dh

    # 否 → 右
    box(d, 1000, 1080, 640, 150,
        "标记为瞬态工况\n只记录原始特征，不报警\n等待进入稳态", fill=LGREY,
        line=GREY, fs=19, bold=False)
    elbow(d, [(cx + dw / 2, dcy), (1320, dcy), (1320, 1080)])
    label(d, cx + dw / 2 + 40, dcy - 30, "否", fs=22, color=GREY, bold=True)

    arrow(d, cx, y, cx, y + 38)
    label(d, cx - 46, y + 20, "是", fs=22, color=GREEN, bold=True)
    y += 38
    y = nbox(y, "与基线比较：多指标偏离度（z 分数）+ 分级迟滞消抖（3 帧确认 / 5 帧恢复）",
             fs=19, h=96)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "状态判定：正常 / 注意 / 报警，并生成维护建议文本", fill=LGREY,
             line=GREY, fs=20, h=96)
    arrow(d, cx, y, cx, y + 30); y += 30
    y = nbox(y, "上报与记录：MQTT 发布（10 s 周期 + 事件即时）+ SD 卡缓存 + "
                "OLED / 声光指示", fs=19, h=96)
    arrow(d, cx, y, cx, y + 26); y += 26
    y = nbox(y, "延时 1 s，进入下一循环（断网自动重连、SD 补传）", fill=LGREEN,
             line=GREEN, fs=19, h=82)

    bx = 56
    elbow(d, [(x, 1380), (bx, 1380), (bx, 268), (x, 268)], color=GREY)
    label(d, bx + 16, 800, "循环采集", fs=20, color=GREY, bold=True, anchor="lm")
    img.save(os.path.join(BASE, "fig3_software.png"))
    trim(os.path.join(BASE, "fig3_software.png"))


# ---------------------------------------------------------------- 图 4 仪表板
def fig_dashboard():
    W, H = 1720, 1040
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([20, 20, W - 20, H - 20], radius=20, fill=(248, 249, 251),
                        outline=GREY, width=3)
    # 顶栏
    d.rounded_rectangle([20, 20, W - 20, 100], radius=20, fill=BLUE)
    d.rectangle([20, 80, W - 20, 100], fill=BLUE)
    label(d, 60, 60, "8.1 振动状态监测器 · Smart Campus Predictive Maintenance",
          fs=24, color=(255, 255, 255), bold=True, anchor="lm")
    label(d, W - 60, 60, "节点：AHU-FAN-01 ｜ 在线 ｜ 中 / EN",
          fs=20, color=(210, 225, 240), anchor="rm")

    # 状态卡
    cards = [("运行状态", "注意", LORANGE, ORANGE),
             ("振动 RMS", "0.86 g", LGREEN, GREEN),
             ("速度有效值", "1.9 mm/s", LORANGE, ORANGE),
             ("表面温度", "46.2 ℃", LGREEN, GREEN)]
    cw = (W - 120 - 3 * 20) / 4
    for i, (k, v, f, l) in enumerate(cards):
        x = 60 + i * (cw + 20)
        box(d, x, 130, cw, 130, f"{k}\n{v}", fill=f, line=l, fs=21, lh=34)

    # 趋势图
    box(d, 60, 290, 1000, 330, "", fill=(255, 255, 255), line=GREY, fs=1)
    label(d, 80, 315, "趋势（近 24 h）：振动 RMS ／ 表面温度", fs=19,
          anchor="lm", bold=True, color=(40, 60, 90))
    d.line([100, 560, 1040, 560], fill=(210, 214, 220), width=2)
    d.line([100, 460, 1040, 460], fill=(210, 214, 220), width=2)
    d.line([100, 360, 1040, 360], fill=(210, 214, 220), width=2)
    d.line([100, 595, 1040, 595], fill=(210, 214, 220), width=2)
    d.line([100, 330, 100, 595], fill=(210, 214, 220), width=2)
    # 曲线：正常 -> 缓慢上升 -> 报警区
    pts = [(100 + i * 47, 560 - (14 * math.sin(i / 2.6) + 0.9 * i * i / 3))
           for i in range(20)]
    d.line(pts, fill=BLUE, width=4, joint="curve")
    d.line([(p[0], p[1] + 26) for p in pts], fill=ORANGE, width=3, joint="curve")
    d.line([100, 430, 1040, 430], fill=RED, width=3)
    label(d, 1032, 418, "报警阈值", fs=17, color=RED, anchor="rm", bold=True)
    d.line([100, 480, 1040, 480], fill=ORANGE, width=2)
    label(d, 1032, 468, "注意阈值", fs=17, color=ORANGE, anchor="rm")

    # 频谱
    box(d, 1090, 290, 570, 330, "", fill=(255, 255, 255), line=GREY, fs=1)
    label(d, 1110, 315, "频谱（1024 点 · Δf 0.98 Hz）", fs=19, anchor="lm",
          bold=True, color=(40, 60, 90))
    base = 595
    peaks = [(1170, 120), (1275, 62), (1385, 96), (1500, 40), (1600, 70)]
    d.line([1110, base, 1640, base], fill=(210, 214, 220), width=3)
    for px, ph in peaks:
        d.rectangle([px - 7, base - ph, px + 7, base], fill=BLUE)
        d.line([px, base - ph, px, base - ph - 8], fill=BLUE, width=2)
    label(d, 1170, base - 148, "基频带峰值", fs=17, color=BLUE, bold=True)
    label(d, 1385, base - 116, "谐波", fs=17, color=BLUE)
    label(d, 1600, base - 90, "谐波", fs=17, color=BLUE)

    # 报警列表
    box(d, 60, 650, 1000, 330, "", fill=(255, 255, 255), line=GREY, fs=1)
    label(d, 80, 678, "报警与建议", fs=19, anchor="lm", bold=True,
          color=(40, 60, 90))
    rows = [("14:22", "注意", "速度有效值 1.9 mm/s 超过注意阈值",
             "建议：检查叶轮积尘与配重"),
            ("13:05", "信息", "表面温度较基线上升 8.2 ℃",
             "建议：检查散热通道与轴承润滑"),
            ("09:40", "正常", "维护后台架复测通过，基线已更新", "—")]
    for i, (t, lv, msg, adv) in enumerate(rows):
        yy = 710 + i * 88
        d.rectangle([90, yy, 1030, yy + 72], fill=(250, 250, 252),
                    outline=(226, 228, 232))
        box(d, 104, yy + 14, 92, 44, lv,
            fill=LRED if lv == "注意" else LGREY,
            line=RED if lv == "注意" else GREY, fs=18, radius=8)
        label(d, 214, yy + 26, t, fs=18, anchor="lm", color=GREY)
        label(d, 286, yy + 26, msg, fs=18, anchor="lm")
        label(d, 286, yy + 52, adv, fs=17, anchor="lm", color=(90, 110, 140))

    # 侧栏
    box(d, 1090, 650, 570, 330, "", fill=(255, 255, 255), line=GREY, fs=1)
    label(d, 1110, 678, "设备与部署", fs=19, anchor="lm", bold=True,
          color=(40, 60, 90))
    items = ["站点：LICHPU 教学楼 AHU-3", "采样：1 kSPS · 帧长 1024",
             "上报：MQTT 10 s / 事件即时", "基线：2026-10-12 开机自学习更新",
             "历史留存：≥ 4 周（SQLite）", "隐私：无摄像头、无个人数据",
             "语言：中文 / English 切换"]
    for i, s in enumerate(items):
        label(d, 1120, 730 + i * 36, "· " + s, fs=18, anchor="lm")
    img.save(os.path.join(BASE, "fig4_dashboard.png"))
    trim(os.path.join(BASE, "fig4_dashboard.png"))


# ---------------------------------------------------------------- 图 5 甘特
def fig_gantt():
    W, H = 1760, 1010
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    left, top, right = 470, 130, W - 60
    colw = (right - left) / 12
    rowh = 52

    tasks = [
        ("问题定义与校园需求（客户访谈）", 0, 1, LGREEN, GREEN),
        ("现有方案对标 · 需求与验收标准", 1, 2, LGREEN, GREEN),
        ("伦理 / 隐私 / 安全筛查", 1, 1, LRED, RED),
        ("系统设计与选型（框图 / BOM）", 2, 1, LBLUE, BLUE),
        ("硬件搭建：传感、供电、台架", 3, 2, LBLUE, BLUE),
        ("固件：采集、特征、网关上报", 3, 3, LORANGE, ORANGE),
        ("集成原型 v1（端到端数据通路）", 4, 2, LORANGE, ORANGE),
        ("仪表板与告警逻辑", 5, 2, LORANGE, ORANGE),
        ("校准（重力翻转 / 参考比对）", 6, 1, LPURPLE, PURPLE),
        ("验证实验：3 类工况变化复现", 7, 2, LPURPLE, PURPLE),
        ("数据分析与部署评估（含 SDG）", 8, 2, LPURPLE, PURPLE),
        ("缺陷修复与设计冻结", 9, 1, LGREY, GREY),
        ("报告、海报、演示准备", 10, 1, LGREY, GREY),
        ("最终评估：演示 + 汇报 + 答辩", 11, 1, LRED, RED),
    ]
    gates = {0: "概念评审", 1: "需求关口", 2: "设计关口", 3: "进度评审 1",
             5: "中期评估", 6: "技术评审", 7: "证据检查点", 9: "设计冻结",
             11: "最终关口"}

    # 表头
    for w in range(12):
        x = left + w * colw
        d.rectangle([x, top - 44, x + colw, top], fill=BLUE)
        label(d, x + colw / 2, top - 22, f"第 {w+1} 周", fs=17,
              color=(255, 255, 255), bold=True)
        d.line([x, top, x, top + len(tasks) * rowh], fill=(216, 220, 226), width=2)
    d.line([right, top, right, top + len(tasks) * rowh], fill=(216, 220, 226), width=2)

    for i, (name, start, dur, f, l) in enumerate(tasks):
        y = top + i * rowh
        label(d, left - 20, y + rowh / 2, name, fs=18, anchor="rm")
        d.line([left, y + rowh, right, y + rowh], fill=(238, 240, 244), width=1)
        bx = left + start * colw + 6
        bwid = dur * colw - 12
        d.rounded_rectangle([bx, y + 12, bx + bwid, y + rowh - 12], radius=8,
                            fill=f, outline=l, width=2)
    # 关口标记
    gy = top + len(tasks) * rowh + 30
    for w, gname in gates.items():
        x = left + w * colw + colw / 2
        d.polygon([(x, gy - 14), (x + 14, gy), (x, gy + 14), (x - 14, gy)],
                  fill=LRED, outline=RED, width=2)
        label(d, x, gy + 42, gname, fs=17, color=RED, bold=True)
    d.line([left, gy, right, gy], fill=(220, 224, 230), width=2)
    img.save(os.path.join(BASE, "fig5_gantt.png"))
    trim(os.path.join(BASE, "fig5_gantt.png"))


if __name__ == "__main__":
    for fn in (fig_arch, fig_datalink, fig_software, fig_dashboard, fig_gantt):
        fn()
        print("saved:", fn.__name__)
