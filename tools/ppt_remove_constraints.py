# -*- coding: utf-8 -*-
"""在现有 PPT 上定向删除「预算 / 成本 / 零焊接」相关表述。

不重新生成整个文件，保留用户的既有修改。
按"段落整段替换"处理，避免 run 被拆分导致替换失败。
"""
import os
import sys

from pptx import Presentation

PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "8.1_振动状态监测器_中期答辩.pptx")

# 关键字：段落含这些词就视为需要处理
KEYWORDS = ("预算", "成本", "¥", "焊接", "已焊针")

# 精确替换表（整段文本 → 新文本）
EXACT = {
    # 第 3 页：约束卡片改为设计目标
    "本课题的约束条件": "本课题的设计目标",
    "预算 ≤ ¥150": "实现可运行的多参数状态监测样机",
    "零焊接（无焊接条件，全部选用已焊针模块）": "所有信号处理与状态判定在设备端完成",
    "实测结果：总支出 ¥119.50，结余 ¥30.50（19.7%）": "系统已可完整运行，具备现场演示条件",

    # 第 4 页：需求卡片
    "成本 ≤ ¥150、零焊接、采样率 ≥ 1 kHz、无线上报、可用性、可维护性、数据不丢失、电气安全、可复制":
        "采样率 ≥ 1 kHz、无线上报、数据不丢失、可用性与可维护性、电气安全、可复制性等 9 条",
    "每条均可量化验证：采样率误差 ≤ 0.5%、静止模长 1.000 ± 0.2 g、报警响应 ≤ 10 s、成本 ≤ ¥150":
        "每条均可量化验证：采样率误差 ≤ 0.5%、静止模长 1.000 ± 0.2 g、报警响应 ≤ 10 s 等",
    "已达成的关键指标：采样率误差 −0.03%（判据 ≤0.5%）、成本 ¥119.50（预算 ¥150）、报警响应 6 s（判据 ≤10 s）":
        "已达成的关键指标：采样率误差 −0.03%（判据 ≤0.5%）、报警响应 6 s（判据 ≤10 s）",

    # 第 6 页：选型理由与原则
    "成本低、I²C 简单、有已焊针版本": "I²C 接口简单、资料丰富、量程满足需求",
    "选型的共同原则：在满足需求前提下优先选「零焊接、低门槛、易复现」的方案，以保证被 LICHPU 复制时不需要专业设备。":
        "选型的共同原则：在满足需求前提下优先选「接口简单、资料丰富、易于复现」的方案，便于在不同环境中复制。",

    # 第 8 页：硬件页标题与说明
    "硬件实现：零焊接的完整样机": "硬件实现：完整样机与安装方式",
    "全部模块选用已焊针版本，用面包板与杜邦线连接，全程无焊接——满足「零焊接」这一硬性约束。":
        "三部分模块通过标准化接口连接，结构清晰、便于拆装与更换。",

    # 第 11 页：进度表
    "三模块全部接入，零焊接": "三模块全部接入并验证",
}


def para_new_text(t):
    t = t.strip()
    if t in EXACT:
        return EXACT[t]
    return None


def fix_paragraph(p):
    """若段落含关键字且命中替换表，则整段替换。返回是否改动。"""
    t = p.text
    new = para_new_text(t)
    if new is None:
        if any(k in t for k in KEYWORDS):
            print("  [未命中] " + t[:60])
        return False
    if new == t:
        return False
    runs = p.runs
    if not runs:
        r = p.add_run()
        r.text = new
        return True
    runs[0].text = new
    for r in runs[1:]:
        r.text = ""
    return True


def walk(shapes, slide_no):
    changed = 0
    for sh in shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                if fix_paragraph(p):
                    changed += 1
                    print("  第 %d 页已改: %s" % (slide_no, p.text[:50]))
        elif sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        if fix_paragraph(p):
                            changed += 1
                            print("  第 %d 页表格已改: %s" % (slide_no, p.text[:50]))
        elif sh.shape_type == 6:  # GROUP
            try:
                changed += walk(sh.shapes, slide_no)
            except Exception:
                pass
    return changed


def main():
    prs = Presentation(PATH)
    total = 0
    for i, s in enumerate(prs.slides, 1):
        total += walk(s.shapes, i)
    prs.save(PATH)
    print("\n共修改 %d 处，已保存" % total)


if __name__ == "__main__":
    main()
