# -*- coding: utf-8 -*-
"""生成《项目 8.1 振动状态监测器 —— 工程报告》。

沿用提案与实施方案的版式（宋体正文 / 黑体标题 / 三线表风格），
所有数据均来自 experiments/ 下的实测记录。
"""
import datetime
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))
FIGDIR = os.path.join(BASE, "figures")
EVDIR = os.path.join(BASE, "evidence")
DATADIR = os.path.join(BASE, "experiments", "data")
OUT = os.path.join(BASE, "8.1_振动状态监测器_工程报告.docx")

DARK = RGBColor(0x1F, 0x3B, 0x63)
BLACK = RGBColor(0x20, 0x20, 0x20)
ACCENT = RGBColor(0xA0, 0x3A, 0x1E)
GRAY = RGBColor(0x50, 0x50, 0x50)

doc = Document()
sec = doc.sections[0]
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.6)


# ---------------------------------------------------------------- 基础排版
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
    # 支持 **粗体** 行内标记：按 ** 切分后交替设置粗体
    segs = text.split("**")
    if len(segs) == 1:
        setfont(p.add_run(text), size=size, bold=bold, ea=ea, color=color)
    else:
        for i, seg in enumerate(segs):
            if not seg:
                continue
            setfont(p.add_run(seg), size=size, bold=(bold or i % 2 == 1),
                    ea=ea, color=color)
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
        p.paragraph_format.line_spacing = 1.26
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


IMG = {
    "arch":    os.path.join(FIGDIR, "figS1_arch.png"),
    "flow":    os.path.join(FIGDIR, "figS2_flow.png"),
    "wiring":  os.path.join(FIGDIR, "figS_wiring.png"),
    "gantt":   os.path.join(FIGDIR, "fig5_gantt.png"),
    "setup":   os.path.join(EVDIR, "setup_overview.jpg"),
    "mount":   os.path.join(EVDIR, "mpu6050_mount.jpg"),
    "probe":   os.path.join(EVDIR, "ds18b20_probe.jpg"),
    "block":   os.path.join(EVDIR, "blockage_test.jpg"),
    "dash_al": os.path.join(EVDIR, "dashboard_live_alarm.png"),
    "dash_du": os.path.join(EVDIR, "dashboard_dual_sensor.png"),
    "spectra": os.path.join(DATADIR, "spectra_compare.png"),
    "history": os.path.join(DATADIR, "history_from_db.png"),
}


def pic(key, caption, width_cm):
    path = IMG[key]
    if not os.path.exists(path):
        para("[缺失图片：%s]" % key, align="center")
        return
    try:
        doc.add_picture(path, width=Cm(width_cm))
    except Exception:
        # 某些 JPEG 缺少 JFIF 标记，python-docx 无法解析，用 PIL 转存为 PNG
        import tempfile
        from PIL import Image
        out = os.path.join(tempfile.gettempdir(),
                           "rpt_img_" + os.path.splitext(os.path.basename(path))[0] + ".png")
        if not os.path.exists(out):
            Image.open(path).convert("RGB").save(out, "PNG")
        doc.add_picture(out, width=Cm(width_cm))
    p = doc.paragraphs[-1]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(8)
    setfont(cp.add_run(caption), size=9, ea="黑体", color=GRAY)


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


# ---------------------------------------------------------------- 页脚页码
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
setfont(fp.add_run("项目 8.1 振动状态监测器 · 工程报告　第 "), size=8)
r = fp.add_run()
f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
r._r.append(f1); r._r.append(it); r._r.append(f2)
setfont(r, size=8)
setfont(fp.add_run(" 页"), size=8)

TODAY = datetime.date.today()


# ================================================================= 封面
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("利莫瑞克大学 × LICHPU　工业评估（15 ECTS）"), size=11.5,
        bold=True, ea="黑体", color=ACCENT)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(14)
setfont(p.add_run("智慧校园生活实验室 · 主题 8 智能设施与预测性维护"), size=11,
        ea="黑体", color=RGBColor(0x55, 0x5B, 0x64))
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
setfont(p.add_run("项目 8.1　振动状态监测器"), size=24, bold=True, ea="黑体",
        color=DARK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(14)
setfont(p.add_run("工程报告"), size=15, bold=True, ea="黑体", color=ACCENT)

table("", ["项目", "内容"],
      [["文件性质", "本项目单节点样机的工程报告：覆盖需求、方案、实现、"
                "测试验证、数据分析与专业实践"],
       ["系统构成", "ESP32 主控 + MPU6050 振动 + DS18B20 表面温度 + microSD 本地存储，"
                "Node-RED 仪表板与 SQLite 持久化"],
       ["核心能力", "边缘完成 1 kSPS 采样、FFT、带限 RMS、基线自学习与迟滞判定；"
                "断网本地缓存与恢复补发；报警时保存原始波形"],
       ["实测范围", "传感器标定、采样率验证、转子不平衡三级工况、风道堵塞三级工况、"
                "端到端联调与断网恢复"],
       ["仓库", "github.com/Zichuan-ALT/vibration-monitor（代码、原始数据、"
              "实验记录与证据照片均已归档）"],
       ["小组", "成员 A　成员 B　成员 C　成员 D"],
       ["日期", TODAY.strftime("%Y-%m-%d")]],
      [Cm(2.8), Cm(13.6)], size=9.5, align_map=["c", "l"])

doc.add_page_break()

# ================================================================= 摘要
h("摘要", 1)
para("本报告介绍一台面向小型旋转设备（以 92 mm 直流风机为被测对象）的振动状态监测样机的"
     "设计、实现与验证。系统以 ESP32 为主控，MPU6050 采集三轴加速度，DS18B20 采集被测体"
     "表面温度，microSD 承担本地记录；设备端完成 1 kSPS 采样、1024 点 FFT、带限 RMS 与"
     "基频带峰值提取、基线自学习与迟滞判定，经 MQTT 上报至 Node-RED 仪表板并写入 SQLite。")
para("关键实测结果：采样率 999.7 Hz，误差 −0.03%（判据 ≤ 0.5%）；带限噪声底 0.0053 g；"
     "0.35 g 配重（约占转子质量 0.4%）使带内 RMS 上升 25.7 倍，1× 转频谱峰稳定在 33~35 Hz；"
     "风道堵塞使表面温度单调上升 +0.62/+0.80 ℃，而振动 RMS 反而下降约 30%，"
     "两类故障的特征方向相反，验证了双通道设计的必要性。")
para("报告同时记录了工程实践中遇到的问题与判断依据，包括一次因变量混淆导致的误判、"
     "以及第三方组件的配置陷阱，并如实说明本方案的局限与后续改进方向。")
para("关键词：状态监测；振动分析；FFT；边缘计算；MQTT；传感器标定", indent=False,
     bold=True, space_after=10)

doc.add_page_break()

# ================================================================= 目录
h("目录", 1)
TOC = [
    ("1  引言", ["1.1 背景与问题", "1.2 项目目标与范围", "1.3 报告结构"]),
    ("2  需求分析", ["2.1 利益相关者与使用场景", "2.2 功能需求",
                 "2.3 非功能需求", "2.4 验收标准与验证方法", "2.5 约束与假设"]),
    ("3  方案设计与选型", ["3.1 总体架构", "3.2 关键设计决策", "3.3 信号处理设计",
                     "3.4 硬件设计", "3.5 软件设计"]),
    ("4  实现", ["4.1 硬件搭建", "4.2 固件实现", "4.3 服务端实现", "4.4 仪表板实现"]),
    ("5  测试与验证", ["5.1 测试策略", "5.2 传感器标定与验证", "5.3 采样与 FFT 验证",
                  "5.4 转子不平衡工况实验", "5.5 频谱分析", "5.6 风道堵塞工况实验",
                  "5.7 系统联调与可靠性验证", "5.8 验收标准符合性"]),
    ("6  数据分析与讨论", ["6.1 检测灵敏度", "6.2 亚线性响应与共振",
                     "6.3 支撑刚度的影响", "6.4 温度通道的时间常数",
                     "6.5 阈值确定方法", "6.6 与设计预期的偏差", "6.7 局限"]),
    ("7  专业实践", ["7.1 版本控制与协作", "7.2 风险管理", "7.3 受控文档",
                 "7.4 调试记录", "7.5 标准与规范"]),
    ("8  负责任的设计", ["8.1 安全", "8.2 隐私", "8.3 可持续性",
                    "8.4 可维护性与可复制性", "8.5 伦理与社会影响"]),
    ("9  UL–LICHPU 复制与推广", []),
    ("10  结论与后续工作", []),
    ("附录", ["附录 A  硬件清单与引脚分配", "附录 B  固件关键参数",
            "附录 C  实验原始数据", "附录 D  小组分工与个人贡献",
            "附录 E  复现步骤", "附录 F  参考文献"]),
]
for title, subs in TOC:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Pt(6)
    setfont(p.add_run(title), size=10.5, bold=True)
    for s in subs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Pt(28)
        p.paragraph_format.line_spacing = 1.2
        setfont(p.add_run(s), size=9.6, color=GRAY)

doc.add_page_break()

# ---------------------------------------------------------------- 正文分章
PARTS = ["report_p1.py", "report_p2.py", "report_p3.py", "report_p4.py"]
for _p in PARTS:
    _path = os.path.join(BASE, _p)
    if not os.path.exists(_path):
        raise SystemExit("缺少分章文件：" + _p)
    with open(_path, encoding="utf-8") as _f:
        exec(compile(_f.read(), _p, "exec"), globals())

doc.save(OUT)
print("已生成：", OUT)
print("段落数：", len(doc.paragraphs), " 表格数：", len(doc.tables))
