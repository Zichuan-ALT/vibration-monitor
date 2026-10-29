# -*- coding: utf-8 -*-
"""把报告正文中与「焊接」相关的表述改为中性说法。"""
import io

RULES = [
    ("report_p1.py",
     'para("本项目的实施受三方面约束：一是硬件条件，实验室不具备焊接条件，"\n'
     '     "因此全部选用已焊排针的成品模块，用面包板与杜邦线连接；"',
     'para("本项目的实施受三方面约束：一是硬件条件，实验室不具备焊接条件，"\n'
     '     "因此全部选用带排针的成品模块，用面包板与杜邦线连接；"'),
    ("report_p2.py",
     '"全部模块均为已焊排针的成品板，用杜邦线直插面包板连接，不需要焊接。"',
     '"全部模块均为带排针的成品板，用杜邦线直插面包板连接，无需任何加工。"'),
    ("report_p2.py",
     '"一是首批 MPU6050 到货为未焊排针的裸板，无法直接插接，后更换为已焊针版本；"',
     '"一是首批 MPU6050 到货为不带排针的裸板，无法直接插接，后更换为带排针的版本；"'),
    ("report_p3.py",
     '[["传感器到货为未焊排针版本", "无法直接插接", "更换为已焊针版本；"\n'
     '        "此后采购前先确认是否已焊针"],',
     '[["传感器到货为不带排针的版本", "无法直接插接", "更换为带排针的版本；"\n'
     '        "此后采购前先确认是否带排针"],'),
    ("report_p4.py",
     '"ESP32-WROOM-32（CH340C，Type-C，已焊排针）"',
     '"ESP32-WROOM-32（CH340C，Type-C，带排针）"'),
    ("report_p4.py",
     '"MPU6050（GY-521，I²C，已焊排针）"',
     '"MPU6050（GY-521，I²C，带排针）"'),
]

for fname, old, new in RULES:
    s = io.open(fname, encoding="utf-8").read()
    if old in s:
        s = s.replace(old, new, 1)
        io.open(fname, "w", encoding="utf-8", newline="\n").write(s)
        print("  已改：", fname)
    else:
        print("  未命中：", fname, "->", old[:36].replace("\n", " "))

print()
for fname, _, _ in RULES:
    s = io.open(fname, encoding="utf-8").read()
    n = s.count("焊接") + s.count("已焊") + s.count("免焊")
    print("  %s 残留 %d 处" % (fname, n))
