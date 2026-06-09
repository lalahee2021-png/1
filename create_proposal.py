from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
section = doc.sections[0]
section.top_margin    = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin   = Cm(3.17)
section.right_margin  = Cm(3.17)

# ── Helper: set paragraph font ────────────────────────────────────────────────
def set_run(run, size=12, bold=False, color=None):
    run.font.name = '宋体'
    run._r.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_para(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT,
             size=12, bold=False, color=None,
             first_line=False, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    if first_line:
        pf.first_line_indent = Pt(size * 2)
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, color=color)
    return p

def add_heading(doc, text, size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                space_before=12, space_after=6):
    return add_para(doc, text, align=align, size=size, bold=bold,
                    space_before=space_before, space_after=space_after)

# ══════════════════════════════════════════════════════════════════════════════
# Title
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '实验四  综合性种植园规划与方案设计', size=16, space_before=0)

# ── 方案简介 ──────────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after  = Pt(6)
r1 = p.add_run('方案简介：')
set_run(r1, size=12, bold=True)
r2 = p.add_run(
    '本方案拟在南方丘陵地区建设一个 30 亩综合性种植园，'
    '选取果树（柑橘、桃树）、蔬菜（番茄、辣椒、豆角）、花卉（月季、百合）'
    '共 3 大类园艺植物，同时辅以食用菌（平菇）栽培区。'
    '园区划分生产区、休闲采摘区、管理服务区三大功能区，'
    '配套建设灌溉、道路、遮阳棚、停车场等基础及配套设施，'
    '实现农业生产与休闲旅游有机融合，打造生态、经济、景观三效合一的现代综合种植园。'
)
set_run(r2, size=12)
p.paragraph_format.first_line_indent = Pt(0)

doc.add_paragraph()  # blank line

# ══════════════════════════════════════════════════════════════════════════════
# 一、种植园功能分区
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '一、种植园功能分区', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

zones = [
    ('（一）生产区（共 22 亩）',
     '位于园区中部及北部平坦地带，是种植园核心区域，主要承担各类园艺植物的规模化生产。'
     '细分为：果树种植区（10 亩，园区北侧）、蔬菜种植区（8 亩，园区中部）、'
     '花卉种植区（3 亩，园区西侧）、食用菌栽培区（1 亩，遮阳大棚内）。'),
    ('（二）休闲采摘区（共 5 亩）',
     '位于园区南部，毗邻入口，视野开阔。设置采摘步道、休憩凉亭、景观节点，'
     '供游客进行果蔬采摘、花卉观赏等休闲活动。区内配置木栈道 200 m、凉亭 2 座。'),
    ('（三）管理服务区（共 3 亩）',
     '位于园区东南角入口处，包括管理用房（100 m²）、农资仓库（60 m²）、'
     '停车场（20 个车位）、公共厕所（30 m²）及游客接待中心（80 m²）。'),
]
for title, body in zones:
    add_para(doc, title, size=12, bold=True, space_before=4, space_after=2)
    add_para(doc, body, size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 二、植物种类、品种配置及栽培面积
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '二、植物种类、品种配置及栽培面积', size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

add_para(doc, '各区植物种类与品种配置详见下表：', size=12, first_line=True)

# Table
table = doc.add_table(rows=1, cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

hdr = table.rows[0].cells
headers = ['类别', '植物种类', '推荐品种', '栽培面积（亩）', '备注']
col_widths = [Cm(2.2), Cm(2.4), Cm(4.6), Cm(2.8), Cm(3.0)]
for i, (cell, h, w) in enumerate(zip(hdr, headers, col_widths)):
    cell.width = w
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    set_run(run, size=11, bold=True)

rows_data = [
    ('果树', '柑橘',   '纽荷尔脐橙、沙糖橘',        '6',  '株距 3 m×4 m'),
    ('果树', '桃树',   '湖景蜜露、霞晖 8 号',         '4',  '株距 3 m×3 m'),
    ('蔬菜', '番茄',   '金鹏 8 号、圣女果（千禧）',  '3',  '大棚+露地轮作'),
    ('蔬菜', '辣椒',   '博辣红牛、线椒 8819',         '3',  '露地种植'),
    ('蔬菜', '豆角',   '之豇 28-2、超级无筋豆',       '2',  '支架栽培'),
    ('花卉', '月季',   '丰花月季混色系列',             '2',  '景观兼采摘'),
    ('花卉', '百合',   '亚洲百合（Landini）',          '1',  '切花生产'),
    ('食用菌', '平菇', '天达 2 号',                    '1',  '遮阳棚袋料栽培'),
]
for row in rows_data:
    cells = table.add_row().cells
    for i, (cell, val, w) in enumerate(zip(cells, row, col_widths)):
        cell.width = w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 2 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(val)
        set_run(run, size=11)

doc.add_paragraph()

add_para(doc,
    '注：果树区合计 10 亩，蔬菜区合计 8 亩，花卉区合计 3 亩，食用菌区 1 亩，'
    '休闲采摘及道路绿化占 5 亩，管理服务区 3 亩，共计 30 亩。',
    size=11, first_line=True, space_before=2)

# ══════════════════════════════════════════════════════════════════════════════
# 三、田间基础设施规划
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '三、田间基础设施规划', size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

infra = [
    ('（一）道路系统',
     '主干道（4 m 宽，硬化混凝土路面）贯穿园区南北，全长约 300 m；'
     '支路（2 m 宽，透水砖铺设）连接各功能分区，总长约 500 m；'
     '田间小路（1.2 m 宽，碎石路面）供农机及人员通行。'),
    ('（二）灌溉系统',
     '果树区采用滴灌系统，管道铺设总长约 800 m，安装滴头 400 个；'
     '蔬菜区采用微喷灌系统，喷头间距 3 m；花卉区采用微滴灌；'
     '全园设主泵房 1 座（30 m²），蓄水池 1 个（容量 200 m³）。'),
    ('（三）排水系统',
     '沿主干道两侧开挖排水沟，深 40 cm、宽 50 cm，内壁水泥抹面；'
     '田间设暗管排水，汇入园区东侧排水总渠，防止积涝。'),
    ('（四）遮阳及保护设施',
     '食用菌区建设钢架遮阳大棚 1 栋（面积 667 m²，遮光率 70%）；'
     '蔬菜区建设连栋塑料大棚 2 栋（各 300 m²），用于反季节蔬菜生产；'
     '果树区设防鸟网覆盖，面积约 4 亩。'),
    ('（五）电力系统',
     '从园区东南角变压器引线，沿主干道铺设三相电缆，'
     '在各功能区设配电箱，满足水泵、大棚补光灯、冷库等用电需求。'),
]
for title, body in infra:
    add_para(doc, title, size=12, bold=True, space_before=4, space_after=2)
    add_para(doc, body, size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 四、园区配套设施
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '四、园区配套设施', size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

support = [
    ('（一）管理服务设施',
     '建设管理用房 1 栋（含办公室、休息室，100 m²）；'
     '农资仓库 1 栋（60 m²，存放农药、化肥、农机具）；'
     '简易冷库 1 间（30 m²，用于采后产品预冷保鲜）。'),
    ('（二）休闲旅游配套',
     '入口景观大门及标识牌 1 套；休憩凉亭 2 座（各 20 m²）；'
     '采摘步道及木栈道 200 m；停车场（硬化，可容纳 20 辆车，面积约 600 m²）；'
     '公共厕所 1 座（30 m²，分男女间）；游客接待中心 1 栋（80 m²，含讲解室、售卖区）。'),
    ('（三）安全与信息化设施',
     '全园安装视频监控摄像头 8 个，覆盖主要区域；'
     '设置水肥一体化智能控制系统 1 套，可远程监控灌溉；'
     '园区内设置科普标识牌 20 块，介绍各植物种类及栽培知识。'),
    ('（四）环保设施',
     '设置垃圾分类收集桶 10 组；建设有机废弃物堆肥池 2 个（各 10 m³）；'
     '污水经简易处理池处理后用于田间灌溉，实现零排放。'),
]
for title, body in support:
    add_para(doc, title, size=12, bold=True, space_before=4, space_after=2)
    add_para(doc, body, size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 五、种植园方案设计图（示意）
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '五、种植园方案设计图', size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

add_para(doc,
    '以下为园区平面布局示意图（文字版），实际图纸可按下图文字描述手绘或电脑制图。',
    size=12, first_line=True)

# ASCII-style layout table (simplified)
layout_table = doc.add_table(rows=7, cols=5)
layout_table.style = 'Table Grid'
layout_table.alignment = WD_TABLE_ALIGNMENT.CENTER

layout_data = [
    ['【北】',        '',              '',           '',          ''],
    ['果树区（柑橘）\n6亩', '果树区（桃树）\n4亩', '蔬菜区（番茄/辣椒）\n6亩', '花卉区\n（月季/百合）\n3亩', '食用菌区\n（遮阳棚）\n1亩'],
    ['',              '← 主干道（4 m 宽，南北贯通）→',  '',  '',  ''],
    ['蔬菜区（豆角）\n2亩', '',   '休闲采摘区\n（含步道/凉亭）\n5亩', '', '管理服务区\n（停车/仓库/接待）\n3亩'],
    ['',              '',              '',           '',          ''],
    ['',   '← 园区入口大门  →',   '',   '',   ''],
    ['【南】',        '',              '',           '',          ''],
]

col_w = [Cm(3.0), Cm(3.5), Cm(3.5), Cm(2.8), Cm(2.2)]
for r_idx, row_data in enumerate(layout_data):
    row = layout_table.rows[r_idx]
    for c_idx, (cell_text, cw) in enumerate(zip(row_data, col_w)):
        cell = row.cells[c_idx]
        cell.width = cw
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(cell_text)
        set_run(run, size=10, bold=(r_idx in [0, 6]))

doc.add_paragraph()

add_para(doc,
    '图注：园区总面积 30 亩，各功能区以主干道为界分隔，支路连接各区；'
    '蓝色线代表灌溉主管，绿色区域代表绿化隔离带（宽 1 m，种植草坪或低矮灌木）；'
    '东侧设排水总渠，雨水统一汇排。',
    size=11, first_line=True, space_before=4)

# ── Save ──────────────────────────────────────────────────────────────────────
out_path = '/home/user/1/实验四_综合性种植园规划与方案设计.docx'
doc.save(out_path)
print(f'Saved: {out_path}')
