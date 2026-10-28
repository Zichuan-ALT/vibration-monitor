# -*- coding: utf-8 -*-
"""把 build_ppt.py 中残留的「预算 / 零焊接 / 已焊针」表述替换掉，
使其与已修改的 PPTX 保持一致。"""
import io
import re

P = "tools/build_ppt.py"
s = io.open(P, encoding="utf-8").read()
before = s

# 1. 第 3 页备注
old1 = ('"这一部分我讲背景：为什么做这件事、以及我们面对的三个约束——"\n'
        '             "预算 150、不能焊接、要能被复制。最后我们实际花了 119.5 元。"')
new1 = ('"这一部分我讲背景：为什么要做这件事、以及本课题的设计目标。"\n'
        '             "重点是说明这类故障在早期都有振动特征的变化，因此可以用低成本方案监测。"')
s = s.replace(old1, new1)

# 2. 第 6 页选型原则
pat2 = re.compile(
    r'para\(tf, "选型的共同原则：在满足需求前提下优先选「零焊接、低门槛、易复现」的方案，"\s*\n'
    r'\s*"以保证被 LICHPU 复制时不需要专业设备。", 14, False, C_GRAY, 0, first=True\)')
new2 = ('para(tf, "选型的共同原则：在满足需求前提下优先选「接口简单、资料丰富、易于复现」"\n'
        '             "的方案，便于在不同环境中复制。", 14, False, C_GRAY, 0, first=True)')
s, n2 = pat2.subn(new2, s)

# 3. 第 8 页绿框
pat3 = re.compile(
    r'para\(tf, "全部模块选用已焊针版本，用面包板与杜邦线连接，全程无焊接"\s*\n'
    r'\s*"——满足「零焊接」这一硬性约束。", 14, True, C_GREEN, 0, first=True\)')
new3 = ('para(tf, "三部分模块通过标准化接口连接，结构清晰、便于拆装与更换。",\n'
        '         14, True, C_GREEN, 0, first=True)')
s, n3 = pat3.subn(new3, s)

# 4. 第 8 页备注
old4 = '"强调零焊接——所有模块买来就是焊好排针的。"'
new4 = '"强调模块通过标准接口连接，方便拆装和替换。"'
n4 = 1 if old4 in s else 0
s = s.replace(old4, new4)

io.open(P, "w", encoding="utf-8", newline="\n").write(s)

print("改动统计：备注1=%s  选型原则=%d  绿框=%d  硬件备注=%d"
      % ("OK" if old1 not in s else "FAIL", n2, n3, n4))
print("文件是否变化:", s != before)

for kw in ("预算", "成本", "¥", "零焊接", "已焊针", "焊接"):
    print("  残留 %-6s %d 处" % (kw, s.count(kw)))
