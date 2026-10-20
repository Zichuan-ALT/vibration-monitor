# -*- coding: utf-8 -*-
"""生成《项目 8.1 振动状态监测器 — 项目提案与需求规格说明》（ESP32 单板方案）。"""
import datetime
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
OUT = os.path.join(BASE, "8.1_振动状态监测器_项目提案与需求规格说明.docx")

DARK = RGBColor(0x1F, 0x3B, 0x63)
BLACK = RGBColor(0x20, 0x20, 0x20)
ACCENT = RGBColor(0xA0, 0x3A, 0x1E)

doc = Document()
sec = doc.sections[0]
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.7)
sec.right_margin = Cm(2.7)


def setfont(run, size=10.5, bold=False, ascii_font="Times New Roman",
            ea="宋体", color=BLACK):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), ascii_font)
    rf.set(qn("w:hAnsi"), ascii_font)
    rf.set(qn("w:eastAsia"), ea)


def fix_style(name, size, ea="宋体", ascii_font="Times New Roman",
              bold=False, color=BLACK):
    st = doc.styles[name]
    st.font.size = Pt(size)
    st.font.bold = bold
    st.font.name = ascii_font
    st.font.color.rgb = color
    rPr = st.element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), ascii_font)
    rf.set(qn("w:hAnsi"), ascii_font)
    rf.set(qn("w:eastAsia"), ea)
    return st


fix_style("Normal", 10.5)
doc.styles["Normal"].paragraph_format.line_spacing = 1.3
doc.styles["Normal"].paragraph_format.space_after = Pt(4)
for lvl, (sz, ea) in {1: (15, "黑体"), 2: (12, "黑体"), 3: (10.8, "黑体")}.items():
    st = fix_style(f"Heading {lvl}", sz, ea=ea, ascii_font="Arial", bold=True,
                   color=DARK)
    st.paragraph_format.space_before = Pt(11 if lvl == 1 else 8)
    st.paragraph_format.space_after = Pt(5)
    st.paragraph_format.line_spacing = 1.15
    st.paragraph_format.keep_with_next = True


def para(text="", *, size=10.5, bold=False, align="justify", ea="宋体",
         color=BLACK, indent=True, space_after=4):
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    setfont(p.add_run(text), size=size, bold=bold, ea=ea, color=color)
    return p


def rich(parts, *, size=10.5, indent=True, space_after=4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    for t, b in parts:
        setfont(p.add_run(t), size=size, bold=b)
    return p


def h(text, level=1):
    p = doc.add_heading("", level=level)
    setfont(p.add_run(text), size={1: 15, 2: 12, 3: 10.8}[level], bold=True,
            ea="黑体", ascii_font="Arial", color=DARK)
    return p


def bullets(items, *, size=10.5):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.27
        p.paragraph_format.left_indent = Pt(20)
        if isinstance(it, str):
            setfont(p.add_run(it), size=size)
        elif isinstance(it, tuple):
            setfont(p.add_run(it[0]), size=size, bold=it[1])
        else:
            for t, b in it:
                setfont(p.add_run(t), size=size, bold=b)


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def set_cell(cell, text, *, size=8.6, bold=False, align="left", ea="宋体"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER}[align]
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.1
    setfont(p.add_run(text), size=size, bold=bold, ea=ea)


def table(caption, headers, rows, widths, *, size=8.6, align_map=None,
          header_fill="DEEBF7"):
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(6)
    cp.paragraph_format.space_after = Pt(2)
    cp.paragraph_format.keep_with_next = True
    if caption:
        setfont(cp.add_run(caption), size=9.5, bold=True, ea="黑体", color=DARK)
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, htxt in enumerate(headers):
        c = t.rows[0].cells[i]
        set_cell(c, htxt, size=size, bold=True, align="center", ea="黑体")
        shade(c, header_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            a = "center" if (align_map and align_map[i] == "c") else "left"
            set_cell(cells[i], str(v), size=size, align=a)
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = widths[i]
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def figure(fname, caption, width_cm):
    doc.add_picture(os.path.join(FIG, fname), width=Cm(width_cm))
    p = doc.paragraphs[-1]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(8)
    setfont(cp.add_run(caption), size=9, ea="黑体",
            color=RGBColor(0x50, 0x50, 0x50))


def code_block(lines):
    for i, ln in enumerate(lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(4) if i == 0 else Pt(0)
        p.paragraph_format.left_indent = Pt(12)
        p.paragraph_format.line_spacing = 1.02
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "F4F6F8")
        pPr.append(shd)
        setfont(p.add_run(ln), size=8.4, ascii_font="Consolas", ea="宋体")
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
setfont(fp.add_run("项目 8.1 振动状态监测器 · 项目提案与需求规格说明　第 "), size=8,
        ea="宋体")
r = fp.add_run()
f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
r._r.append(f1); r._r.append(it); r._r.append(f2)
setfont(r, size=8)
setfont(fp.add_run(" 页"), size=8, ea="宋体")

TODAY = datetime.date.today()

# ================================================================= 封面
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("利莫瑞克大学 × LICHPU　工业评估（15 ECTS）"), size=11.5,
        bold=True, ea="黑体", color=ACCENT)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
setfont(p.add_run("智慧校园生活实验室 · 主题 8 智能设施与预测性维护"), size=11,
        ea="黑体", color=RGBColor(0x55, 0x5B, 0x64))
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("项目 8.1　振动状态监测器"), size=25, bold=True, ea="黑体",
        color=DARK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
setfont(p.add_run("Vibration Condition Monitor"), size=13, bold=True, ea="黑体",
        color=DARK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
setfont(p.add_run("项目提案与需求规格说明（评估组成 A）"), size=13, bold=True,
        ea="黑体", color=ACCENT)

table("", ["项目", "内容"],
      [["项目编号 / 名称", "8.1　振动状态监测器（Vibration Condition Monitor）"],
       ["主题领域", "主题 8：智能设施与预测性维护"],
       ["指示性范围", "在安全的演示用电机/风扇上测量振动，并识别运行状况的变化"],
       ["技术路线", "ESP32-WROOM-32 单板 + MPU6050（加速度 + 板载温度）+ DS18B20 + "
                "Node-RED 仪表板"],
       ["本文件对应关口", "第 1 周 概念评审　→　第 2 周 需求关口　→　"
                   "第 3 周 设计关口"],
       ["评估组成", "A. 项目提案与需求（10%）——本文件为主要提交物"],
       ["配套文件", "8.1_振动状态监测器_实施方案.docx（施工与调试手册）\n"
                "8.1_受控文档台账.xlsx（需求、物料、风险、责任矩阵与计划）"],
       ["团队与角色", "P1 项目/客户协调员　P2 硬件与集成　P3 软件与数据　"
                 "P4 测试/验证与文档\n（姓名与学号：＿＿＿＿＿＿＿＿＿＿＿＿）"],
       ["版本 / 日期", f"V3.0　{TODAY.strftime('%Y-%m-%d')}"]],
      [Cm(3.2), Cm(12.8)], size=9.5, align_map=["c", "l"])

# ================================================================= 摘要
h("执行摘要", 1)
para("本文件针对智慧校园生活实验室项目 8.1“振动状态监测器”提出完整的问题定义、"
     "需求规格与概念设计。校园内大量小型旋转设备（空调机组风机、通风风机、循环水泵）"
     "长期连续运行，其轴承磨损、叶轮积尘、转子不平衡、底座松动等劣化过程会先表现为"
     "振动与温度的缓慢漂移，而现有维护方式以定期巡检和事后报修为主，既看不到趋势，"
     "也无法在故障前给出预警。", indent=True)
para("我们提出的方案是一台面向该类设备的低成本、开放式状态监测节点：以振动加速度与"
     "两路温度共三路数字传感输入为基础，在边缘侧完成 1024 点 FFT、速度有效值与基频带"
     "特征提取，通过开机基线自学习与多指标偏离度判据，输出“正常 / 注意 / 报警”三级状态"
     "与可执行的维护建议；经 MQTT 上报至笔记本上的 Node-RED 仪表板，实现实时监视、"
     "历史趋势与报警闭环。系统全程工作于 5 V USB 安全低压演示台架，不接触市电，"
     "不采集任何个人数据，服务端复用现有笔记本，无额外硬件采购。", indent=True)
para("本文件共 4 幅插图、15 张表，给出 11 项功能性需求、9 项非功能性需求、"
     "13 条可测量的验收标准、12 周计划与关口对应关系、责任矩阵、物料与耗材预算"
     "（实际支出 ¥119.50，结余 ¥30.50）以及 13 项风险登记条目，"
     "并专门讨论了在 UL 与 LICHPU 两地复制部署的可行性与差异。"
     "具体的施工步骤、接线与代码要点见配套的《实施方案》。", indent=True)

# ================================================================= 一
h("一、问题定义与校园需求", 1)
h("1.1 问题陈述", 2)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Pt(16)
p.paragraph_format.right_indent = Pt(16)
p.paragraph_format.space_before = Pt(4)
pPr = p._p.get_or_add_pPr()
shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear")
shd.set(qn("w:fill"), "FDF2E9"); pPr.append(shd)
setfont(p.add_run("校园设施团队目前只能依靠定期巡检和故障报修来维护空调风机、"
                  "通风风机与循环水泵；我们要为这些设备装上一个能连续“听诊”的"
                  "低成本节点，让它用振动和温度数据在故障发生前就指出设备正在变坏。"),
        size=11, bold=True, ea="黑体")
para("受益者是校园设施/物业维护团队。可衡量成果：在人工巡检之前发现设备状态变化，"
     "把维护动作从“坏了再修”前移到“按状态维护”，从而减少非计划停机与不必要的更换。",
     indent=True)

h("1.2 受益者与利益相关者", 2)
table("表 1　利益相关者与其核心诉求", ["角色", "关注点", "本项目如何回应"],
      [["设施/物业维护工程师", "哪些设备在变坏、先修哪个、要不要停机",
        "三级状态 + 可执行维护建议 + 优先级排序"],
       ["能源与可持续发展管理人员", "劣化设备能耗升高、非计划更换产生额外碳排",
        "振动与温度趋势可作为能效与寿命的代理指标（见第七章）"],
       ["实验室/教学楼使用者", "空调与通风设备可靠、噪声与舒适度稳定",
        "提前发现不平衡与积尘，避免噪声上升与风量下降"],
       ["模块教员（客户与技术评审）", "需求可测量、工程深度足够、证据充分",
        "本文件的需求与验收标准、第九章的测试与校准计划"],
       ["IT / 网络运维", "新增设备不增加网络与安全负担",
        "标准 MQTT、受限带宽、无个人数据、可限域部署"]],
      [Cm(3.4), Cm(5.6), Cm(7.0)], size=8.6, align_map=["l", "l", "l"])

h("1.3 候选部署地点", 2)
table("表 2　候选部署地点与选择理由", ["地点", "设备", "选择理由", "状态"],
      [["LICHPU 教学楼 AHU-3", "空调机组离心风机（约 1 500 r/min）",
        "有独立检修空间，可加装传感器，安全低压", "首选试点"],
       ["LICHPU 实验楼通风风机", "轴流风机", "连续运行、负载稳定，适合长期趋势验证",
        "备选"],
       ["UL 校园同类机组", "等效风机 / 循环水泵", "用于跨国复制验证，比较两地设备与环境差异",
        "第二阶段"],
       ["项目自建演示台架", "5 V USB 小风扇",
        "可控可复现：可加配重、可调松动、可遮挡风道；不接触市电", "主要验证平台"]],
      [Cm(3.4), Cm(4.2), Cm(5.6), Cm(2.8)], size=8.6,
      align_map=["l", "l", "l", "c"])

h("1.4 成功衡量标准", 2)
table("表 3　项目成功衡量标准", ["维度", "衡量标准", "目标"],
      [["技术", "能否在人工巡检之前识别出人工制造的工况变化",
        "3 类工况变化中至少正确识别 2 类"],
       ["数据质量", "连续运行与数据完整性", "≥ 72 h 连续运行，数据丢失率 ≤ 1 %"],
       ["可用性", "维护人员能否看懂并据此行动",
        "≥ 3 名非项目成员能在无人讲解下读懂状态卡与建议"],
       ["成本", "单节点成本", "≤ ¥150（含耗材与备件）"],
       ["可复制", "能否在 UL 与 LICHPU 两地复制", "给出双地采购、环境适配与部署路径（第八章）"],
       ["负责任", "隐私、安全与伦理", "无摄像头、无个人数据；全程 5 V USB 供电、免焊接装配"]],
      [Cm(2.4), Cm(6.8), Cm(6.8)], size=8.6, align_map=["l", "l", "l"])

# ================================================================= 二
h("二、现有解决方案与实践现状", 1)
para("为了说明本项目的设计空间，我们对商用状态监测产品、低成本无线振动传感器以及"
     "开源/教学方案进行了对标，重点比较传感能力、数据开放性、部署成本与是否具备"
     "边缘处理能力。", indent=True)
table("表 4　现有方案对标分析", ["方案类别", "代表产品/做法", "能力", "不足（相对本项目需求）"],
      [["便携式测振仪", "Fluke 805 FC、SKF QuickCollect",
        "点检精度高，内置 ISO 10816 判据，工业认可",
        "人工点检、非连续；无趋势与历史；数据接口封闭；单价约 1 万元以上"],
       ["工业无线振动变送器", "Emerson AMS 9420、ifm VVB001、Banner QM30VT",
        "振动+温度一体，防护等级高，支持长期在线监测",
        "单价 2 000 ~ 8 000 元/测点；协议与网关封闭，需厂商平台；校园预算难以规模部署"],
       ["开源/教学方案", "ESP32 + MPU6050 + Grafana、TinyML 振动异常检测示例",
        "成本极低（< ¥100），资料多，适合快速原型",
        "多为单一传感器；缺少校准与不确定度说明；阈值靠经验设定；缺少验证证据"],
       ["楼宇自控系统（BMS）", "既有空调机组自带振动报警", "与设备集成度高",
        "通常仅有开关量报警，无频谱与趋势；不对外提供原始数据"]],
      [Cm(2.6), Cm(4.2), Cm(4.2), Cm(5.0)], size=8.4,
      align_map=["l", "l", "l", "l"])
rich([("对标结论与设计空间：", True),
      ("商用方案准但贵、封闭，适合“诊断”不适合“密集布点”；开源方案便宜但工程深度不足，"
       "通常止步于“采集 + 画图”。本项目的设计空间因此明确为："
       "以三路数字传感输入的融合判据、边缘特征提取与基线自学习、"
       "以及开放可互操作的数据模型，在单节点 ≤ ¥150 的成本下，"
       "提供接近工业方案的“状态变化识别”能力与完整的证据链。", False)])

# ================================================================= 三
h("三、需求与验收标准", 1)
h("3.1 功能性需求", 2)
table("表 5　功能性需求（优先级：M 必须 / S 应该 / C 可以）",
      ["编号", "需求描述", "优先级", "验证方法"],
      [["FR-01", "采集振动（I²C，1 kSPS，帧长 1024 点）与两路温度，共三路数据", "M",
        "数据记录与仪表板检查"],
       ["FR-02", "计算时域特征：RMS、峰值、速度有效值（10 ~ 250 Hz 积分）", "M",
        "注入已知信号比对"],
       ["FR-03", "完成 1024 点 FFT，给出主频与基频带（15 ~ 35 Hz）内峰值", "M",
        "音箱振膜 + 100 Hz 正弦验证"],
       ["FR-04", "支持开机基线自学习，并给出相对基线的偏离度", "M", "与离线计算比对"],
       ["FR-05", "输出“正常 / 注意 / 报警”三级状态与维护建议", "M",
        "人工制造工况变化验证"],
       ["FR-06", "通过 MQTT 上报特征值，周期 ≤ 10 s，事件即时上报", "M",
        "抓包与端到端计时"],
       ["FR-07", "断网时本地缓存，网络恢复后补传，不丢数据", "S", "拔网线测试"],
       ["FR-08", "提供中英双语仪表板，含实时值、历史趋势与频谱", "M",
        "界面走查 + 用户测试"],
       ["FR-09", "提供本地指示（板载 LED + 有源蜂鸣器模块），无网络时仍可判读", "S",
        "断网走查"],
       ["FR-10", "记录报警事件（含触发指标与时间戳）并可导出", "S", "导出并核对"],
       ["FR-11", "支持按键重新学习基线与静音", "C", "现场操作验证"]],
      [Cm(1.5), Cm(9.2), Cm(1.5), Cm(3.8)], size=8.6,
      align_map=["c", "l", "c", "l"])

h("3.2 非功能性需求", 2)
table("表 6　非功能性需求", ["编号", "类别", "需求描述"],
      [["NFR-01", "安全", "全部电路 ≤ 5 V 直流（USB 供电），不接触市电；转动部件加防护罩"],
       ["NFR-02", "隐私", "不使用摄像头或麦克风，不采集任何可识别个人的数据"],
       ["NFR-03", "可靠性", "连续运行 ≥ 72 h 无重启；看门狗、掉电恢复、断网自愈"],
       ["NFR-04", "可维护性", "传感器可现场更换且免重新编译；基线可一键重学"],
       ["NFR-05", "成本", "单节点 ≤ ¥150（含耗材与备件）"],
       ["NFR-06", "可访问性", "仪表板中英双语；状态用文字 + 图标 + 颜色三重编码；字号可调"],
       ["NFR-07", "互操作性", "数据模型带版本号并有文档，可被第三方平台订阅复用"],
       ["NFR-08", "装配", "全程无需焊接，仅使用杜邦线与免焊端子"],
       ["NFR-09", "环境", "外壳防护 ≥ IP54，工作温度 0 ~ 50 ℃"]],
      [Cm(1.7), Cm(1.9), Cm(12.4)], size=8.6, align_map=["c", "c", "l"])

h("3.3 验收标准", 2)
para("验收标准均可测量、可复现，并与评估组成 D（测试、验证与数据分析，占 20%）"
     "直接对应。", indent=True)
table("表 7　验收标准与验证方法", ["编号", "验收判据", "目标值", "验证方法"],
      [["AC-01", "振动采样率准确度", "误差 ≤ 0.5 %", "连续采集 1 s 统计实际点数"],
       ["AC-02", "频谱频率准确度", "误差 ≤ 1 % 或 ≤ Δf", "音箱振膜 + 100 Hz 正弦"],
       ["AC-03", "幅值准确度（比对）", "偏差 ≤ 10 %", "与手持测振仪同点同时比对"],
       ["AC-04", "温度测量准确度", "± 0.5 ℃", "冰水浴与温水比对"],
       ["AC-05", "工况变化识别", "3 类中至少 2 类在 60 s 内检出并分级正确",
        "不平衡 / 松动 / 负载变化对照实验"],
       ["AC-06", "误报控制", "正常工况 2 h 内假阳性 ≤ 1 次", "正常工况长时运行"],
       ["AC-07", "报警端到端延迟", "≤ 5 s", "注入越限信号并计时"],
       ["AC-08", "连续运行能力", "≥ 72 h 无重启，丢数率 ≤ 1 %", "不间断运行并统计记录"],
       ["AC-09", "断网恢复", "≤ 30 s 内自动重连并补传", "拔插网线测试"],
       ["AC-10", "仪表板刷新与历史", "刷新 ≤ 5 s；历史可查 ≥ 4 周", "走查 + 数据库查询"],
       ["AC-11", "单节点成本", "≤ ¥150", "按表 12 物料清单核算"],
       ["AC-12", "重复性", "变异系数 ≤ 3 %", "同一工况重复测量 10 次"],
       ["AC-13", "免焊接装配", "全部连接为插拔或免焊端子", "装配过程走查与记录"]],
      [Cm(1.5), Cm(4.6), Cm(5.0), Cm(4.9)], size=8.5,
      align_map=["c", "l", "l", "l"])

h("3.4 范围界定", 2)
bullets([
    [("本项目做：", True), ("面向小型旋转设备的连续状态监测与“变化识别”，"
                        "包括两路传感采集、边缘特征提取、三级判别、报警与趋势可视化。",
                        False)],
    [("本项目不做：", True), ("不接触市电设备与配电柜；不采集音视频与个人数据；"
                          "不给出轴承剩余寿命（RUL）的精确预测；不做电网侧电能质量分析；"
                          "不在未经批准的生产设备上长期部署。", False)],
])

# ================================================================= 四
h("四、伦理、隐私与安全筛查", 1)
para("按手册第 13 章要求，在需求阶段完成筛查。结论：本项目采用保护隐私、无摄像头的"
     "物理量传感方案，全部工作在 5 V USB 安全低压演示台架上进行，", indent=True)
rich([("不需要", True),
      ("教员就“市电、个人数据、摄像头/视觉、健康相关测量、无人机、危险环境”作出"
       "特别批准；若后续扩展到真实运行设备取电，则须重新提交批准申请。", False)],
     indent=False)
table("表 8　伦理、隐私与安全风险筛查", ["关注项", "是否存在", "评估与控制措施", "剩余风险"],
      [["市电与电气安全", "否", "全系统 ≤ 5 V USB 适配器供电；不拆改配电柜；"
                          "台架加防护罩", "低"],
       ["个人数据", "否", "不采集姓名、图像、位置或行为数据；仅设备振动与温度", "无"],
       ["摄像头/视觉", "否", "方案不含任何成像器件", "无"],
       ["健康相关测量", "否", "不涉及人体生理参数", "无"],
       ["设备损坏与误停机", "是", "仅监测不控制，不接入设备控制系统；报警仅为建议性质",
        "低"],
       ["数据安全", "是", "MQTT 使用独立账号与最小权限；数据库限域访问；"
                     "保留期 ≤ 4 周", "低"],
       ["学术诚信", "是", "记录第三方库与 AI 工具使用；原始数据与处理脚本入库（Git）",
        "低"],
       ["第三方场地许可", "是", "试点部署前取得设施团队书面同意；先台架后现场", "低"]],
      [Cm(2.8), Cm(1.8), Cm(8.6), Cm(2.4)], size=8.5,
      align_map=["l", "c", "l", "c"])

# ================================================================= 五
h("五、系统架构与概念设计", 1)
h("5.1 总体架构", 2)
para("系统按感知层、节点层、网络层、应用层四层组织，核心设计原则是"
     "“边缘先算出结论、网络只传结论与必要证据”。", indent=True)
figure("figS1_arch.png", "图 1　系统总体架构", 15.8)

h("5.2 传感与数据采集", 2)
para("手册要求“通常至少两个有意义的数据输入/传感器”。本方案采用两路数字接口传感，"
     "并使二者在判据上互补：振动反映机械劣化的即时证据，温度反映负载与润滑的慢变量。",
     indent=True)
bullets([
    [("① 振动 + 板载温度 — MPU6050 模块（I²C，±4 g，ODR 1 kSPS，板载温度）。", True),
     ("安装于风机外壳靠近轴承的位置，贴紧刚性安装，覆盖 10 ~ 260 Hz，"
      "足以覆盖 1 500 r/min 风机的 1× ~ 10× 转频范围与常见轴承特征频率。", False)],
    [("② 温度 — DS18B20（1-Wire，−55 ~ +125 ℃，±0.5 ℃，12 bit）。", True),
     ("贴附于同一测点附近的外壳表面，用于捕捉润滑劣化、负载上升与散热变差"
      "引起的温升趋势。", False)],
])
rich([("为什么两路都要：", True),
      ("风机积尘与转子不平衡会立刻抬升振动，而散热不良或轴承润滑劣化往往先表现为温度"
       "上升、振动后起。单一指标容易漏判：只看振动会漏掉热相关的早期劣化，"
       "只看温度则对结构性缺陷迟钝。两路联合判据是本方案工程深度的主要来源。", False)])

h("5.3 边缘处理与决策逻辑", 2)
para("固件在本地完成全部特征计算与判别，只上传结论与少量证据数据，"
     "满足手册“有意义的本地处理、过滤、融合、检测或决策逻辑——"
     "不仅仅是转发原始数值”的要求。主流程见图 2。", indent=True)
figure("figS2_flow.png", "图 2　固件主流程（含基线自学习与三级判别）", 10.4)
rich([("基线自学习：", True),
      ("开机后连续采集 60 s，统计正常工况下振动 RMS 的均值 μ 与标准差 σ，"
       "自动生成注意阈值 μ+3σ 与报警阈值 μ+6σ；温度采用相对基线的增量阈值。"
       "基线只在“连续 30 min 处于正常状态”时缓慢更新，避免把已经劣化的状态"
       "学成新的正常。", False)])
rich([("多指标偏离度与迟滞：", True),
      ("对振动 RMS 与表面温度分别计算偏离度，任一越限即升级状态等级；"
       "报警需连续 3 帧确认，恢复需连续 5 帧，形成迟滞以抑制阈值附近的抖动。", False)])
rich([("基频带处理：", True),
      ("本方案未配置转速传感器，因此不对单次测量的 1× 做精确跟踪，"
       "而是取 15 ~ 35 Hz（对应 900 ~ 2100 r/min）这一“基频带”内的最大谱峰"
       "及其频率作为指标。这样在转速小范围波动时结果仍然稳定，"
       "代价是无法做严格的阶次归一化——该局限在第九章中如实说明。", False)])

h("5.4 通信与数据模型", 2)
bullets([
    "链路：ESP32 内置 Wi-Fi → MQTT Broker（与仪表板同机）→ SQLite 时序存储 → "
    "Node-RED 仪表板。",
    "报文：JSON，带 schema 版本号；周期状态 10 s 一条，事件即时上报。",
    "带宽：单节点平均 < 10 kbps，满足校园网络与多节点扩展。",
    "互操作性：字段名、单位与取值范围在附录 B 中完整定义，"
    "第三方平台（含未来数字孪生）可直接订阅。",
])
code_block([
    '主题：smartcampus/vib/{node_id}/state   （周期，retained）',
    '      smartcampus/vib/{node_id}/event   （事件，即时）',
    '      smartcampus/vib/{node_id}/cmd     （下行：阈值下发 / 基线重学）',
])

h("5.5 仪表板与用户界面", 2)
para("仪表板面向设施维护人员而非工程师，因此首屏必须回答三个问题："
     "这台设备现在好不好、比过去差了多少、我该做什么。线框设计见图 3。", indent=True)
figure("fig4_dashboard.png", "图 3　中英双语仪表板线框：状态卡 · 趋势 · 频谱 · 报警与建议",
       15.8)
bullets([
    "实时区：状态指示灯 + 振动 RMS、速度有效值、表面温度三张状态卡；",
    "趋势区：近 24 h 曲线并叠加注意/报警阈值线；支持 7 天与 4 周视图；",
    "频谱区：1024 点幅值谱，标注基频带；",
    "报警与建议区：分级列表，每条报警附一条可执行的维护建议；",
    "可访问性：中英双语切换、状态用“文字 + 图标 + 颜色”三重编码、字号可调。",
])

h("5.6 演示台架与安全设计", 2)
bullets([
    "台架：5 V USB 风扇固定在木板上，配重片制造不平衡、可调松动底座制造结构松动、"
    "纸板遮挡风道制造负载变化；全部由 5 V USB 适配器供电，转动部件加防护罩。",
    "可复现性：每种工况变化都有明确的施加方式与量级记录（配重质量、螺栓圈数、"
    "遮挡比例），保证实验可重复。",
    "安全：不接触市电；传感器与线缆固定牢靠，避免转动部件卷入。",
])

h("5.7 关键技术选型", 2)
table("表 9　关键技术选型与理由", ["环节", "选用", "备选", "选型理由"],
      [["主控", "ESP32-WROOM-32 单板",
        "STM32F407、ESP32-S3",
        "一块板同时完成采集、计算与联网，省掉模块间通信；240 MHz 双核 + FPU "
        "足以完成 1024 点 FFT"],
       ["振动 + 温度", "MPU6050 模块（GY-521）",
        "ICM-42688-P、ADXL355、IEPE 压电加速度计",
        "一颗芯片同时给三轴加速度与板载温度，I²C 仅 4 线；省掉独立温度芯片与全部焊接工序"],
       ["温度传感器", "DS18B20", "SHT31（I²C）、NTC + 分压",
        "数字接口抗干扰、单线布线简单、精度满足趋势监测"],
       ["服务端", "Node-RED + SQLite（笔记本）",
        "Mosquitto + InfluxDB + Grafana（三容器）",
        "一个进程同时充当 Broker、历史库与仪表板，零硬件采购、开发量最小"],
       ["仪表板", "Node-RED dashboard", "自研 Flask + Chart.js",
        "拖拽式开发，双语与阈值线配置容易，实时与历史兼备"]],
      [Cm(2.2), Cm(3.6), Cm(4.0), Cm(6.2)], size=8.4,
      align_map=["l", "l", "l", "l"])

# ================================================================= 六
h("六、项目计划与资源", 1)
h("6.1 十二周计划与关口", 2)
figure("fig5_gantt.png", "图 4　十二周计划、任务分布与正式关口", 15.8)
table("表 10　关口对应与所需证据", ["周次", "关口", "本项目的证据物"],
      [["第 1 周", "概念评审", "本文件第一、二章：问题陈述、利益相关者、地点、成功标准、对标"],
       ["第 2 周", "需求关口", "本文件第三章与第四章：需求、验收标准、伦理/隐私/安全筛查"],
       ["第 3 周", "设计关口", "本文件第五章、表 12 物料清单、表 13 风险登记册"],
       ["第 4 周", "进度评审 1", "振动与温度传感器的单元验证记录"],
       ["第 5 周", "—", "频谱验证记录（100 Hz 正弦比对）"],
       ["第 6 周", "中期评估", "可演示的端到端系统（采集 → 判别 → 上报 → 仪表板）"],
       ["第 7 周", "技术评审", "校准报告与不确定度预算"],
       ["第 8 周", "证据检查点", "三类工况变化的对照实验数据集"],
       ["第 9 周", "进度评审 2", "数据分析报告、可持续性与复制性评估"],
       ["第 10 周", "设计冻结", "缺陷关闭清单、72 h 连续运行记录"],
       ["第 11 周", "草案就绪", "报告草稿、海报/幻灯片、仓库与文档完成"],
       ["第 12 周", "最终关口", "原型演示、团队汇报、个人答辩、最终报告与工程日志"]],
      [Cm(1.6), Cm(2.5), Cm(11.9)], size=8.5, align_map=["c", "l", "l"])

h("6.2 责任矩阵", 2)
table("表 11　责任矩阵（R 主责 / S 支持 / C 审核 / I 知会）",
      ["交付物", "P1 项目/客户", "P2 硬件/集成", "P3 软件/数据", "P4 测试/文档"],
      [["问题定义与需求规格", "R", "S", "S", "C"],
       ["系统设计与选型", "C", "R", "S", "S"],
       ["硬件搭建与台架", "I", "R", "S", "S"],
       ["固件与边缘算法", "I", "S", "R", "C"],
       ["通信与仪表板", "I", "S", "R", "S"],
       ["校准与测试计划", "C", "S", "S", "R"],
       ["验证实验执行与数据分析", "S", "S", "C", "R"],
       ["风险登记册与项目记录", "R", "S", "S", "C"],
       ["报告、演示与个人日志", "C", "S", "S", "R"]],
      [Cm(5.0), Cm(2.7), Cm(2.7), Cm(2.7), Cm(2.5)], size=8.5,
      align_map=["l", "c", "c", "c", "c"])

h("6.3 物料与耗材预算", 2)
para("下表列出主要器件与耗材分类；逐项清单（14 项规格、数量与单价）见配套的"
     "《实施方案》表 3 ~ 表 7 与受控文档台账。", indent=True)
table("表 12　物料与耗材预算（预算上限 ¥150）", ["项目", "内容", "金额 ¥"],
      [["主要器件", "ESP32 开发板 ×2（CH340C，已焊排针）、MPU-6050 三轴加速度模块、"
                "DS18B20 温度模块 + 防水探头、microSD 卡座与 32 GB Class10 A1 卡、"
                "有源蜂鸣器模块、5 V 9225 演示风扇", "112.50"],
       ["连接件", "TYPE-C 数据线、USB 公头转杜邦线；杜邦线与面包板由学校提供", "7.00"],
       ["结构与电源", "外壳、底板、固定与配重；电源使用手边现有 USB 充电器", "0"],
       ["实际支出合计", "4 人分摊约 ¥30 / 人", "119.50"],
       ["预算上限", "满足 NFR-05（≤ ¥150）", "150"],
       ["结余", "余量 ¥30.50（约 20 %）", "30.50"],
       ["装配工序", "全程免焊接：已焊排针模块 + 杜邦线直插 + 面包板电源轨", "—"],
       ["服务端", "复用现有笔记本，零硬件采购", "0"]],
      [Cm(4.4), Cm(8.6), Cm(2.8)], size=8.6, align_map=["l", "l", "c"])

h("6.4 风险登记册", 2)
table("表 13　风险登记册（等级 = 概率 × 影响，1 低 / 2 中 / 3 高）",
      ["ID", "风险描述", "类别", "等级", "应对措施", "责任人"],
      [       ["R-01", "传感器安装方式不当导致高频幅值失真", "技术", "4",
        "模块贴紧外壳刚性安装，标定记录安装方式并做频响修正", "P2"],
       ["R-02", "无标准振动台，幅值无法绝对溯源", "资源", "6",
        "重力翻转静态标定 + 手持测振仪比对 + 报告不确定度", "P4"],
       ["R-03", "Wi-Fi 覆盖不足或丢包", "基础设施", "4",
        "本地 SD 先存后传；必要时用手机热点", "P3"],
       ["R-04", "MPU6050 带宽仅 260 Hz，限制可诊断的故障类型", "技术", "9",
        "确认目标故障频率均 < 130 Hz 并写入报告；预留升级到 ICM-42688-P 的路径", "P3"],
       ["R-05", "台架转速过低，基频带内无可观测峰值", "技术", "4",
        "选用 1200 r/min 以上风扇；必要时下移基频带并重新标定", "P2"],
       ["R-06", "到货与预期不符（排针未焊、随附线缆为纯充电线）", "采购", "6",
        "下单明确勾选“已焊排针”；到货当天验货，不合格立即换货；另备一根数据线", "P2"],
       ["R-07", "采样节拍被 Wi-Fi 中断，频谱失真", "技术", "6",
        "采样任务绑定 core 1 并设高优先级；采样用定时器而非 delay", "P3"],
       ["R-08", "MQTT 报文超长导致静默发送失败", "技术", "6",
        "首行设置 setBufferSize(1024)；增加发送返回值检查", "P3"],
       ["R-09", "器件缺货或交期过长", "采购", "6",
        "提前 2 周下单；ESP32 与 MPU6050 各备一件", "P2"],
       ["R-10", "团队时间冲突（其他模块考核）", "项目", "6",
        "周会 + WBS 分解到人；关键路径任务提前一周启动", "P1"],
       ["R-11", "microSD 卡损坏导致历史丢失", "技术", "4",
        "使用 Class10 卡；每日校验落盘；重要数据每周导出到仓库", "P3"],
       ["R-12", "Node-RED 或笔记本不可用", "基础设施", "4",
        "节点本地闭环运行；恢复后补传", "P3"],
       ["R-13", "校准不确定度导致结论不可靠", "验证", "4",
        "给出不确定度预算，结论附置信区间，避免过度断言", "P4"]],
      [Cm(0.9), Cm(4.4), Cm(1.5), Cm(1.0), Cm(6.1), Cm(1.1)], size=8.0,
      align_map=["c", "l", "c", "c", "l", "c"])

# ================================================================= 七
h("七、可持续性与可持续发展目标（SDG）贡献", 1)
para("预测性维护的价值不止于少修一次设备：劣化的风机因不平衡与积尘而效率下降，"
     "在同样风量下需要更多电能；非计划更换则带来额外的材料与运输碳排。"
     "本项目通过延长设备寿命、减少非计划停机与避免不必要更换，贡献于以下目标。",
     indent=True)
table("表 14　SDG 贡献与可测量证据", ["SDG", "本项目的贡献", "可测量证据"],
      [["SDG 7 可负担的清洁能源", "减少劣化设备的多余电耗",
        "同一风量下不同工况的功率 / 占空比代理指标对比"],
       ["SDG 9 产业、创新与基础设施", "以低成本方案提升基础设施维护水平",
        "单节点成本与检出率（单位成本效益）"],
       ["SDG 11 可持续城市与社区", "提升校园建筑设备可靠性与舒适度",
        "非计划停机次数、噪声与风量记录"],
       ["SDG 12 负责任的消费与生产", "按状态维护，减少不必要的部件更换",
        "试验期内避免的更换次数与物料估算"],
       ["SDG 13 气候行动", "延长设备寿命，减少隐含碳排",
        "寿命延长估算 × 部件碳因子（引用文献值）"]],
      [Cm(3.4), Cm(6.4), Cm(6.2)], size=8.5, align_map=["l", "l", "l"])

# ================================================================= 八
h("八、UL–LICHPU 跨国复制与部署路径", 1)
para("手册要求讨论在 UL 与 LICHPU 两地复制的可能性，包括采购、服务可用性、语言与"
     "环境差异。本设计从器件、供电、环境、网络、设备与法规六个维度做了对照。",
     indent=True)
table("表 15　UL 与 LICHPU 复制差异与对策", ["维度", "差异", "对策"],
      [["采购与成本", "中国采购实际支出 ¥119.50；爱尔兰本地分销约 €55 ~ €75",
        "集中采购后随身携带（注意锂电池航空限制）；批量采购摊薄运费与关税"],
       ["供电", "爱尔兰 230 V / 中国 220 V，均为 50 Hz",
        "本方案使用 USB 5 V 供电，只需符合当地认证的 USB 充电器（CE / CCC），节点侧接口不变"],
       ["气候与环境", "爱尔兰湿冷多雨易凝露；中国夏季高温高湿",
        "IP54 外壳 + 透气防水阀；接线端子防潮；温度阈值用与环境温差而非绝对温度"],
       ["网络", "两地校园 Wi-Fi 均可用 2.4 GHz 信道 1 ~ 13",
        "节点不硬编码信道；MQTT Broker 在两地各部署或使用云端"],
       ["设备与基线", "两地风机品牌、叶片数与转速不同",
        "基线自学习使阈值自适应，无需按品牌重新标定；仅需一次现场基线采集"],
       ["语言与可访问性", "两地语言不同",
        "仪表板中英双语；状态用文字 + 图标 + 颜色三重编码"],
       ["法规与数据保护", "爱尔兰适用 GDPR；中国适用 PIPL",
        "不采集个人数据，仅聚合设备物理量，满足数据最小化；数据限域存储、保留期 ≤ 4 周"]],
      [Cm(2.4), Cm(5.6), Cm(7.6)], size=8.5, align_map=["l", "l", "l"])
rich([("部署路线：", True),
      ("第一阶段在 LICHPU 自建台架完成全部验证；第二阶段在一台真实空调风机上试运行 4 周，"
       "比较台架与现场的基线差异；第三阶段把同一节点与同一数据模型搬到 UL 的等效设备上，"
       "用同一套仪表板对比两地数据，验证“一次设计、两地部署”的可行性。", False)])

# ================================================================= 九
h("九、结论与下一步", 1)
para("本文件把选题 8.1“在安全的演示用电机/风扇上测量振动，并识别运行状况的变化”"
     "转化为一套可测量、可验证、可复制的工程需求：振动与两路温度共三路数字传感输入"
     "满足“至少两个有意义数据输入”的要求；边缘侧完成 FFT、速度有效值、"
     "基线自学习与三级判别，满足“有意义的本地处理而非转发原始数值”；"
     "MQTT + SQLite + 双语仪表板满足连通性与可视化要求；"
     "全部工作在 12 V 安全低压台架、不使用摄像头与个人数据，"
     "满足伦理、隐私与安全要求。", indent=True)
para("本方案在取舍上也保持了诚实：不配置转速传感器，因而无法做严格的阶次归一化，"
     "判据仅在接近额定转速时可靠；不做包络解调，对轴承早期点蚀灵敏度有限，"
     "主要识别不平衡、积尘、松动与负载变化四类——恰好覆盖校园风机最常见的故障。"
     "这两点在第九章的局限说明与答辩中都会如实陈述，"
     "并给出了相应的增量升级路径（见《实施方案》第十三章）。", indent=True)
para("下一步（对应第 2 ~ 4 周）：", indent=True)
bullets([
    "在概念评审上确认问题陈述与受益者，取得设备与场地的书面许可；",
    "完成需求关口签署，冻结 FR / NFR 与验收标准基线；",
    "下单采购（优先处理长交期器件），开始台架搭建与传感器单元验证；",
    "确认实验室能否借到手持测振仪——这是 AC-03 的前提条件；",
    "同步搭建 Node-RED 最小链路，用模拟数据先把仪表板跑通。",
])

# ================================================================= 附录
doc.add_page_break()
h("附录 A　需求追溯矩阵", 1)
table("", ["需求", "验收标准", "测试阶段", "关口证据"],
      [["FR-01 两路传感采集", "AC-01 / AC-04", "单元 + 集成", "第 4 周"],
       ["FR-02 时域特征", "AC-03 / AC-12", "单元", "第 5 周"],
       ["FR-03 FFT 与基频带", "AC-02", "单元", "第 5 周"],
       ["FR-04 基线自学习与偏离度", "AC-06", "系统", "第 6 周"],
       ["FR-05 三级状态与建议", "AC-05 / AC-06", "系统 + 验证实验", "第 8 周"],
       ["FR-06 MQTT 上报与延迟", "AC-07", "集成", "第 6 周"],
       ["FR-07 断网缓存与补传", "AC-09", "系统", "第 10 周"],
       ["FR-08 双语仪表板", "AC-10", "用户评估", "第 6 周"],
       ["FR-09 本地指示", "AC-09", "系统", "第 10 周"],
       ["FR-10 事件记录与导出", "AC-10", "系统", "第 9 周"],
       ["FR-11 按键交互", "—", "现场操作", "第 10 周"],
       ["NFR-01 电气安全", "—", "走查", "第 3 周"],
       ["NFR-03 可靠性", "AC-08", "系统", "第 10 周"],
       ["NFR-05 成本", "AC-11", "核算", "第 3 周"],
       ["NFR-06 可访问性", "—", "用户评估", "第 6 周"]],
      [Cm(4.6), Cm(4.2), Cm(3.4), Cm(3.4)], size=8.6,
      align_map=["l", "l", "l", "c"])

h("附录 B　MQTT 主题与 JSON 数据模型", 1)
para("数据模型带版本号，字段名、单位与取值范围如下；第三方平台可直接订阅复用。",
     indent=True)
code_block([
    '{',
    '  "schema": "vib.node.v1",',
    '  "node_id": "AHU-FAN-01",',
    '  "site": "LICHPU/CB-AHU3",',
    '  "ts": "2026-10-13T14:22:03Z",',
    '  "fw": "1.0.0",',
    '  "state": "attention",          // normal | attention | alarm',
    '  "score": 2.7,                  // 相对基线的偏离度',
    '  "vib": {',
    '    "rms_g": 0.86,               // 振动加速度有效值 [g]',
    '    "peak_g": 3.10,',
    '    "v_rms_mmps": 1.90,          // 速度有效值 [mm/s]，10-250 Hz',
    '    "band_peak_hz": 24.6,        // 基频带 15-35 Hz 内峰值频率',
    '    "band_peak_g": 0.42          // 该峰值幅值',
    '  },',
    '  "temp": { "surface_c": 46.2, "delta_c": 6.4 },',
    '  "baseline": { "rms_g": 0.52, "sigma_g": 0.041, "temp_c": 39.8,',
    '                "updated": "2026-10-12T09:10:00Z" },',
    '  "quality": { "sensor_ok": true, "samples": 1024, "fs_hz": 1000 },',
    '  "advice": "inspect_impeller_balance"',
    '}',
])

h("附录 C　个人工程日志模板", 1)
table("", ["字段", "填写要求"],
      [["日期 / 时长", "本次投入的时间"],
       ["本周目标", "与 WBS 对应的任务编号"],
       ["完成的工作", "具体动作与产出（附提交哈希或文件路径）"],
       ["产生的证据", "数据、截图、测试记录、评审反馈编号"],
       ["遇到的阻塞", "问题描述与升级路径"],
       ["工程决策", "做了什么决定、依据是什么、放弃了什么备选方案"],
       ["反思", "如果重做会怎样改进"],
       ["下周计划", "任务与截止日"]],
      [Cm(4.0), Cm(12.0)], size=8.8, align_map=["l", "l"])

h("附录 D　待模块负责人确认的字段", 1)
table("", ["项目", "待填写内容"],
      [["模块代码 / 名称", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["学年 / 学期", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["项目起止日期", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["提案 / 中期 / 最终报告截止日期", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["演示与答辩日期", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["提交平台", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["经批准的实验室与硬件安排", "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿"],
       ["是否批准使用真实运行设备", "□ 仅限台架　□ 允许低风险现场试点"],
       ["实验室是否提供手持测振仪", "□ 提供　□ 不提供（需调整 AC-03 口径）"]],
      [Cm(6.4), Cm(9.6)], size=9, align_map=["l", "l"])

doc.save(OUT)
print("saved:", OUT)
