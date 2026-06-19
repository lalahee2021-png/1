#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成朱顶红切块繁殖与栽培管理实践报告（Word格式）"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ──────────────────────────────────────────────
# 页面设置  A4
# ──────────────────────────────────────────────
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.top_margin    = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin   = Cm(3.17)
section.right_margin  = Cm(3.17)

# ──────────────────────────────────────────────
# 全局字体默认：宋体 12pt
# ──────────────────────────────────────────────
style = doc.styles['Normal']
style.font.name = '宋体'
style.font.size = Pt(12)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# ──────────────────────────────────────────────
# 帮助函数
# ──────────────────────────────────────────────
def set_font(run, name='宋体', size=12, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    if color:
        run.font.color.rgb = RGBColor(*color)

def heading(doc, text, level=1, font_size=16, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    set_font(run, name='黑体', size=font_size, bold=bold)
    return p

def body(doc, text, indent=0, space_before=0, space_after=4, first_line=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if first_line:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_font(run, name='宋体', size=12)
    return p

def bold_body(doc, label, content, indent=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    r1 = p.add_run(label)
    set_font(r1, name='宋体', size=12, bold=True)
    r2 = p.add_run(content)
    set_font(r2, name='宋体', size=12)
    return p

def add_table_border(table):
    """给表格加边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def set_cell_text(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=11):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    set_font(run, name='宋体', size=size, bold=bold)

# ══════════════════════════════════════════════
# 封面
# ══════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()
doc.add_paragraph()

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_after = Pt(10)
r = title_p.add_run('朱顶红切块繁殖与栽培管理技术')
set_font(r, name='黑体', size=22, bold=True)

subtitle_p = doc.add_paragraph()
subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle_p.paragraph_format.space_after = Pt(6)
r2 = subtitle_p.add_run('实  践  报  告')
set_font(r2, name='黑体', size=18, bold=True)

doc.add_paragraph()
doc.add_paragraph()
doc.add_paragraph()

for label, value in [
    ('实践地点：', '植物园温室大棚'),
    ('小组名称：', '第三实践小组'),
    ('小组成员：', '张明、李华、王芳'),
    ('指导教师：', '陈老师'),
    ('实践时间：', '2026年4月10日 — 2026年6月19日'),
    ('提交日期：', '2026年6月19日'),
]:
    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_p.paragraph_format.space_after = Pt(8)
    r_l = info_p.add_run(label)
    set_font(r_l, name='宋体', size=14, bold=True)
    r_v = info_p.add_run(value)
    set_font(r_v, name='宋体', size=14)

doc.add_page_break()

# ══════════════════════════════════════════════
# 一、实践目的
# ══════════════════════════════════════════════
heading(doc, '一、实践目的', font_size=16)

body(doc,
     '朱顶红（Hippeastrum hybridum）为石蒜科朱顶红属多年生球根花卉，以花大色艳、'
     '形态端庄著称，广泛应用于室内观赏与园林景观配置。其鳞茎储存大量营养物质，分生'
     '能力较强，适合通过切块方式实现快速、规模化繁殖。',
     first_line=True)

body(doc,
     '本次实践依托校内植物园温室大棚，以朱顶红切块繁殖为核心操作内容，结合系统性'
     '栽培管理，旨在达成以下学习目标：',
     first_line=True)

for item in [
    '（1）熟练掌握朱顶红鳞茎切块繁殖的操作规范与关键技术要点；',
    '（2）深入理解朱顶红各生长阶段的发育规律，并能进行科学的环境调控；',
    '（3）系统学习朱顶红日常栽培管理方法，包括水肥调控与病虫害综合防治；',
    '（4）扩展了解朱顶红播种、分球、组织培养等其他繁殖方式，开阔知识视野。',
]:
    body(doc, item, indent=0.5)

body(doc,
     '通过理论与实践相结合，培养观察记录能力、团队协作意识和科学分析思维，'
     '为今后从事花卉生产与园艺管理工作奠定基础。',
     first_line=True)

# ══════════════════════════════════════════════
# 二、实践内容
# ══════════════════════════════════════════════
heading(doc, '二、实践内容', font_size=16)

# ── 2.1 前期准备
heading(doc, '（一）前期准备', font_size=14)

bold_body(doc, '1. 场地准备　', '实践地点为植物园中温室大棚，温湿度可调控，光照充足，适合切块繁殖与后续养护。')

body(doc, '2. 分组与分工', indent=0)
body(doc, '本小组共3人，分工如下：张明负责鳞茎切割与消毒操作；李华负责基质准备、扦插及日常养护记录；王芳负责拍摄记录、数据统计与报告撰写。小组每日碰头，共同讨论植株生长动态。', first_line=True)

body(doc, '3. 材料与工具清单', indent=0)

# 材料清单表
tbl1 = doc.add_table(rows=1, cols=3)
tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl1)
hdr = tbl1.rows[0].cells
set_cell_text(hdr[0], '类别', bold=True)
set_cell_text(hdr[1], '名称', bold=True)
set_cell_text(hdr[2], '规格 / 用量', bold=True)

materials = [
    ('试验材料', '朱顶红成熟母球', '直径8-10cm，6个'),
    ('试验材料', '朱顶红子球',     '直径3-4cm，10个'),
    ('消毒用品', '75%酒精',         '500mL'),
    ('消毒用品', '0.1%多菌灵溶液', '配制2L'),
    ('消毒用品', '草木灰/硫磺粉',  '适量'),
    ('栽培基质', '腐叶土:园土:河沙=5:3:2', '拌匀，约15kg'),
    ('器具',    '枝剪、菜刀',       '各1把，消毒备用'),
    ('器具',    '花盆（直径20cm）', '20个'),
    ('器具',    '育苗盘、喷壶',     '各2个'),
    ('其他',    '保鲜膜、标签牌',   '若干'),
]
for row_data in materials:
    row = tbl1.add_row().cells
    for i, txt in enumerate(row_data):
        align = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT
        set_cell_text(row[i], txt, align=align)

doc.add_paragraph()

# ── 2.2 切块繁殖操作
heading(doc, '（二）朱顶红切块繁殖操作', font_size=14)

heading(doc, '1. 最佳繁殖时间选择', font_size=13)
body(doc,
     '本次实践于2026年4月10日启动，正值春季3—4月，朱顶红鳞茎从冬季休眠状态'
     '逐渐复苏，生命活力旺盛，愈伤组织形成快，切块繁殖成功率高。若错过春季，'
     '秋季9—10月气温回落也适宜操作；夏季高温（>35℃）和冬季低温（<10℃）时均'
     '应避免切割，以防伤口感染或鳞茎腐烂。',
     first_line=True)

heading(doc, '2. 种球处理与消毒', font_size=13)
for step in [
    '（1）将选定的成熟母球置于通风处放置1-2天，使外层鳞片略微干燥，便于切割操作；',
    '（2）徒手剥去外层枯黄干瘪的鳞片，露出内层白色健康鳞片，同时检查是否存在病斑或虫口；',
    '（3）用75%酒精棉球擦拭鳞茎表面及切割工具（菜刀、枝剪），晾干备用；',
    '（4）切除鳞茎底部老化根系及腐烂组织，保留完整健康的鳞茎盘（根盘），这是切块能否萌发新球的关键结构。',
]:
    body(doc, step, indent=0.5)

heading(doc, '3. 切割方法与操作要点', font_size=13)
body(doc,
     '本次实践同时采用十字切割法与楔形切割法进行对比：',
     first_line=True)
bold_body(doc, '十字切割法：', '将鳞茎纵向从顶端切至基部，再横向切一刀，得到4个等份切块。每块保留约1/4的鳞茎盘及2-3个鳞片层，操作简便，适合初学者。')
bold_body(doc, '楔形切割法：', '以鳞茎中心为轴，纵向切成8-16块，每块呈楔形，须确保带有至少1-2个芽眼及部分鳞茎盘。此法繁殖系数更高，但操作精度要求较高。')
body(doc,
     '切割过程中，刀具须随时保持消毒，避免交叉感染；切面要光滑整齐，勿撕裂鳞片'
     '组织；每块宽度不小于1cm，以确保储存足够的营养物质支撑新芽萌发。',
     first_line=True)

heading(doc, '4. 伤口处理与晾干', font_size=13)
for step in [
    '（1）切割完成后，将所有切块浸入0.1%多菌灵溶液中消毒10分钟，捞出后平铺于干净纸张上沥干；',
    '（2）待表面水分挥发后，在每个切口均匀涂抹草木灰（或硫磺粉），起到收敛伤口、抑制杂菌的作用；',
    '（3）将切块放置在阴凉通风处静置1-2天，待切口表面形成薄层愈伤组织后再行扦插，可显著降低腐烂率。',
]:
    body(doc, step, indent=0.5)

heading(doc, '5. 扦插基质准备与扦插', font_size=13)
body(doc,
     '将河沙与珍珠岩按1:1混合，使用0.1%高锰酸钾溶液浇透消毒，晾干至半湿状态'
     '后装入育苗盘（厚度约5cm）。扦插时将切块芽眼朝上，平置于基质表面，切勿埋入'
     '土中，否则切口积水易腐烂。切块间距5-10cm，操作完成后用保鲜膜覆盖育苗盘，'
     '以保湿保温。',
     first_line=True)

heading(doc, '6. 扦插后养护管理', font_size=13)

# 养护参数表
tbl2 = doc.add_table(rows=1, cols=3)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl2)
hdr2 = tbl2.rows[0].cells
set_cell_text(hdr2[0], '管理要素', bold=True)
set_cell_text(hdr2[1], '目标参数', bold=True)
set_cell_text(hdr2[2], '管理措施', bold=True)

manage_data = [
    ('温度', '25—30℃', '温室恒温控制，不足时开启加热设备'),
    ('湿度', '空气湿度70—80%\n基质保持湿润', '每日用喷壶少量喷水，避免积水'),
    ('光照', '遮阴，避免直射光', '催芽期保持黑暗或弱散射光'),
    ('施肥', '扦插后1个月起', '喷施稀薄叶面肥（N:P:K=1:1:1），每10天1次'),
]
for row_data in manage_data:
    row = tbl2.add_row().cells
    set_cell_text(row[0], row_data[0], bold=True)
    set_cell_text(row[1], row_data[1])
    set_cell_text(row[2], row_data[2], align=WD_ALIGN_PARAGRAPH.LEFT)

doc.add_paragraph()

heading(doc, '7. 移栽定植', font_size=13)
body(doc,
     '经过约2-3个月的培养，鳞片块间出现直径0.5-1cm的小种球。当小种球长出2-3条白色'
     '嫩根后，即可移栽至疏松透气的基质（腐叶土:园土:河沙=5:3:2）中，种植深度为小种球'
     '直径的1.5倍，置于散射光环境中缓苗，保持土壤微润。',
     first_line=True)

# ── 2.3 栽培管理
heading(doc, '（三）朱顶红栽培管理', font_size=14)

heading(doc, '1. 土壤配制与上盆', font_size=13)
body(doc,
     '将腐叶土、园土、河沙按5:3:2的体积比充分混合，并拌入适量腐熟有机肥（每盆约10g）。'
     '选用直径20cm的花盆，盆底铺2-3cm碎砖粒以利排水。种植时令种球露出土面约1/3，'
     '浇透定根水后置于光线良好处。',
     first_line=True)

heading(doc, '2. 光照管理', font_size=13)
body(doc,
     '营养生长期每日保证不少于6小时直射光照，有利于叶片充分光合、鳞茎积累养分；'
     '夏季温度超过32℃时在正午前后适当遮阴（遮光率约30%），防止灼叶；冬季休眠期'
     '可移至室内明亮散射光处。',
     first_line=True)

heading(doc, '3. 温度管理', font_size=13)
body(doc,
     '生长适温18—25℃；花期前后适当提高至20—28℃可促进花梗抽出。冬季休眠期控温'
     '5—10℃，使鳞茎充分休眠并积累花芽分化所需物质，但需避免0℃以下低温冻害。',
     first_line=True)

heading(doc, '4. 水肥管理', font_size=13)
body(doc,
     '生长旺盛期（4—9月）保持盆土湿润，以"见干见湿"为原则，避免盆底积水；'
     '每15天施一次稀薄饼肥水（1:20稀释），花前花后各增施一次磷钾肥（如磷酸二氢钾'
     '0.2%溶液）；秋末逐渐减少浇水，待叶片枯黄后停水，进入休眠。',
     first_line=True)

heading(doc, '5. 病虫害防治', font_size=13)
body(doc, '（1）常见病害及防治：')

# 病害表
tbl3 = doc.add_table(rows=1, cols=4)
tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl3)
hdr3 = tbl3.rows[0].cells
for i, h in enumerate(['病害名称', '主要症状', '发生条件', '防治措施']):
    set_cell_text(hdr3[i], h, bold=True)

diseases = [
    ('红斑病', '叶片产生红褐色椭圆形病斑，边缘紫红', '高温高湿，通风不良', '发病初期喷50%多菌灵800倍液，每7天1次，连喷3次'),
    ('叶斑病', '叶面出现黄褐色斑点，后期病斑扩大融合', '降雨频繁，植株过密', '喷施75%百菌清600倍液；合理密植，加强通风'),
    ('细菌性\n软腐病', '鳞茎及叶基出现水渍状腐烂，伴异味', '土壤积水，伤口感染', '及时清除病株，土壤用石灰消毒；避免积水'),
]
for row_data in diseases:
    row = tbl3.add_row().cells
    for i, txt in enumerate(row_data):
        align = WD_ALIGN_PARAGRAPH.LEFT if i in [1,3] else WD_ALIGN_PARAGRAPH.CENTER
        set_cell_text(row[i], txt, align=align, size=10)

doc.add_paragraph()
body(doc, '（2）常见虫害及防治：')

# 虫害表
tbl4 = doc.add_table(rows=1, cols=3)
tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl4)
hdr4 = tbl4.rows[0].cells
for i, h in enumerate(['虫害名称', '危害特征', '防治措施']):
    set_cell_text(hdr4[i], h, bold=True)

pests = [
    ('石蒜夜蛾', '幼虫钻入鳞茎内部蛀食，造成空洞腐烂', '发现时用敌敌畏1000倍液浇灌基部；人工捕杀成虫'),
    ('介壳虫',   '固着于叶片、鳞茎吸食汁液，导致叶黄萎蔫', '喷施40%速扑杀乳油1000倍液；人工刷除虫体'),
    ('红蜘蛛',   '叶片背面结网，出现灰白色细点，严重时叶片焦枯', '喷施73%克螨特2000倍液，注意喷及叶背'),
]
for row_data in pests:
    row = tbl4.add_row().cells
    for i, txt in enumerate(row_data):
        align = WD_ALIGN_PARAGRAPH.LEFT if i != 0 else WD_ALIGN_PARAGRAPH.CENTER
        set_cell_text(row[i], txt, align=align, size=10)

doc.add_paragraph()

# ── 2.4 观察记录
heading(doc, '（四）实验观察记录', font_size=14)

heading(doc, '1. 切块繁殖观察数据', font_size=13)
body(doc,
     '本次实践共切割母球6个，采用十字法得切块24块，楔形法得切块48块，合计72块。'
     '以下为扦插后60天的统计结果：',
     first_line=True)

# 繁殖结果统计表
tbl5 = doc.add_table(rows=1, cols=5)
tbl5.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl5)
hdr5 = tbl5.rows[0].cells
for i, h in enumerate(['切割方法', '总切块数', '成功出球数', '腐烂数', '出球率（%）']):
    set_cell_text(hdr5[i], h, bold=True)

stats = [
    ('十字切割法', '24', '19', '5', '79.2'),
    ('楔形切割法', '48', '35', '13', '72.9'),
    ('合计', '72', '54', '18', '75.0'),
]
for row_data in stats:
    row = tbl5.add_row().cells
    bold = row_data[0] == '合计'
    for i, txt in enumerate(row_data):
        set_cell_text(row[i], txt, bold=bold)

doc.add_paragraph()
p_note = doc.add_paragraph('注：出球率 = 成功形成小种球的切块数 / 总切块数 × 100%；腐烂率 = 腐烂切块数 / 总切块数 × 100% = 25.0%。')
p_note.paragraph_format.left_indent = Cm(1)
for run in p_note.runs:
    set_font(run, size=10)

heading(doc, '2. 栽培管理记录（鳞茎生长数据）', font_size=13)
body(doc,
     '随机抽取10盆朱顶红（子球上盆后60天），测量以下指标，结果如下：',
     first_line=True)

# 生长数据表
tbl6 = doc.add_table(rows=1, cols=6)
tbl6.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl6)
hdr6 = tbl6.rows[0].cells
for i, h in enumerate(['盆号', '鳞茎直径(cm)', '鳞茎盘径(cm)', '单球重(g)', '叶片长(cm)', '叶片宽(cm)']):
    set_cell_text(hdr6[i], h, bold=True, size=10)

growth_data = [
    ('1','5.2','2.1','38.5','22.3','3.2'),
    ('2','4.8','1.9','32.1','19.8','2.9'),
    ('3','5.5','2.3','41.2','24.1','3.4'),
    ('4','5.0','2.0','36.8','21.5','3.1'),
    ('5','4.6','1.8','30.5','18.9','2.8'),
    ('6','5.3','2.2','39.6','23.0','3.3'),
    ('7','5.1','2.0','37.2','22.0','3.2'),
    ('8','4.9','1.9','34.8','20.5','3.0'),
    ('9','5.4','2.2','40.3','23.5','3.3'),
    ('10','5.0','2.1','36.0','21.8','3.1'),
    ('均值','5.08','2.05','36.70','21.74','3.13'),
]
for i, row_data in enumerate(growth_data):
    row = tbl6.add_row().cells
    bold = row_data[0] == '均值'
    for j, txt in enumerate(row_data):
        set_cell_text(row[j], txt, bold=bold, size=10)

doc.add_paragraph()

# ══════════════════════════════════════════════
# 三、实践结论与分析
# ══════════════════════════════════════════════
heading(doc, '三、实践结论与分析', font_size=16)

heading(doc, '（一）朱顶红栽培管理技术总结', font_size=14)

body(doc,
     '通过本次实践，总结出朱顶红高效栽培管理的以下核心要点：',
     first_line=True)

for item in [
    '1. 基质配制：以腐叶土:园土:河沙=5:3:2为最优配比，此比例兼顾保水性与透气性，'
    '鳞茎根系发育良好，无积水烂根现象。实践证明，单纯使用园土或沙土均会导致植株'
    '生长迟缓或徒长。',

    '2. 水肥调控：遵循"生长期见干见湿、休眠期控水断肥"的原则，每15天补充一次稀薄'
    '饼肥水，花期前增施磷钾肥，显著促进了花梗粗壮和花色鲜艳。过量施氮导致叶片徒长'
    '而影响开花的问题在初期实验盆中有所体现，后期调整后改善明显。',

    '3. 光温管理：朱顶红对光照需求较高，日照不足6小时时叶片变薄、色淡；夏季午间'
    '温度超过32℃时，适度遮阴有效防止了灼叶。冬季保持5—10℃低温休眠，为翌年'
    '开花积累了充足的花芽分化物质。',

    '4. 病虫害防治：本次实践中红斑病发生较为普遍（占总盆数的20%），主要原因是梅雨季节'
    '通风不畅。喷施多菌灵后3周内病情有效控制。介壳虫少量发生，人工刮除辅以速扑杀'
    '药剂取得良好效果。实践证明，预防性措施（通风、控湿、消毒）的成本远低于事后治疗。',
]:
    body(doc, item, first_line=True, space_after=6)

heading(doc, '（二）切块繁殖操作方法归纳与技术总结', font_size=14)

body(doc,
     '朱顶红切块繁殖技术可分为"准备—消毒—切割—愈伤—扦插—养护—移栽"七个阶段，'
     '各阶段技术要点归纳如下：',
     first_line=True)

# 技术总结表
tbl7 = doc.add_table(rows=1, cols=3)
tbl7.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_border(tbl7)
hdr7 = tbl7.rows[0].cells
for i, h in enumerate(['操作阶段', '关键技术要点', '常见问题与对策']):
    set_cell_text(hdr7[i], h, bold=True)

tech_summary = [
    ('种球准备',
     '选直径8-10cm健壮母球；\n提前1-2周停水使鳞茎干燥',
     '种球带病斑→提前用高锰酸钾浸泡消毒'),
    ('器具消毒',
     '刀具用75%酒精擦拭；\n基质用高锰酸钾溶液浇透',
     '消毒不彻底→切块感染率上升，需严格执行'),
    ('切割操作',
     '每块须带根盘组织；\n宽度≥1cm；刀面光滑',
     '切口撕裂→换锋利刀具，一刀到位'),
    ('伤口处理',
     '多菌灵浸泡10min；\n涂草木灰；晾1-2天',
     '晾干不足→切口霉变，需延长晾干时间'),
    ('扦插',
     '芽眼朝上，平置基质表面；\n间距5-10cm；覆保鲜膜',
     '切块倾斜→接触不良，及时调整角度'),
    ('养护',
     '温度25-30℃；湿度70-80%；\n黑暗催芽；1个月后叶面施肥',
     '温度不稳定→出球时间延长，需加强温控'),
    ('移栽',
     '出2-3条根后移栽；\n种植深度=小球径×1.5倍',
     '过早移栽→根系弱，成活率低；宜耐心等待'),
]
for row_data in tech_summary:
    row = tbl7.add_row().cells
    set_cell_text(row[0], row_data[0], bold=True)
    set_cell_text(row[1], row_data[1], align=WD_ALIGN_PARAGRAPH.LEFT, size=10)
    set_cell_text(row[2], row_data[2], align=WD_ALIGN_PARAGRAPH.LEFT, size=10)

doc.add_paragraph()

body(doc,
     '本次实践综合出球率为75.0%，十字切割法（79.2%）略优于楔形切割法（72.9%），'
     '主要原因在于十字法切块保留的根盘组织面积更大、伤口数更少，感染概率相对较低。'
     '楔形法虽繁殖系数更高（单球可得8-16块），但对操作精度要求更高，建议初学者先掌握'
     '十字法再过渡至楔形法，以提高综合繁殖效率。',
     first_line=True)

body(doc,
     '与其他繁殖方式相比：播种法后代性状分离，从播种到开花需3-5年；分球法操作最简单'
     '但繁殖系数仅1:2；组织培养法繁殖系数最高但成本昂贵。切块繁殖法在繁殖系数、成本、'
     '操作难度和遗传稳定性间取得了较好的平衡，适合中等规模的种苗生产。',
     first_line=True)

# ══════════════════════════════════════════════
# 四、实践体会
# ══════════════════════════════════════════════
heading(doc, '四、实践体会', font_size=16)

heading(doc, '（一）个人心得', font_size=14)

body(doc,
     '这是我第一次系统性地参与植物球根繁殖的全程操作，从最初对"每个切块必须带根盘"'
     '这一要点一知半解，到亲眼看到小种球在鳞片间破壳而出的那一刻，内心的震撼难以言表。'
     '书本上枯燥的参数——温度25-30℃、湿度70-80%、黑暗催芽——在实际操作中变成了'
     '每天细心观察的具体行动：摸盆土干湿、测温室温度、及时揭膜换气。这让我深刻体会到，'
     '植物栽培是一门需要用手、用眼、用心共同完成的学问，而不仅仅是照着配方执行。',
     first_line=True)

body(doc,
     '在病虫害防治环节，我学会了"预防重于治疗"的核心理念。梅雨季节红斑病的爆发让'
     '整个小组措手不及，事后复盘才意识到是前期通风管理疏漏所致。这次教训让我明白，'
     '日常管理细节的积累才是保障植株健康的根本，药剂只能是最后的补救手段。',
     first_line=True)

heading(doc, '（二）小组不足之处', font_size=14)
for item in [
    '1. 数据记录频率不够稳定：前两周记录较为认真，中期因课程安排紧张出现了连续4天未测量的情况，导致部分生长曲线数据不连续，影响了分析的精准度；',
    '2. 照片记录不够系统：初期未统一拍摄角度和光线条件，导致不同阶段照片对比效果不佳，建议后续实践固定拍摄位点和时间（如每周同一时间、同一角度）；',
    '3. 对照试验设计欠缺：两种切割方法虽同步进行，但未设置基质类型、施肥浓度等单因素对照组，导致无法量化分析各因素对出球率的独立影响；',
    '4. 分工协调有待提升：前期分工过细导致成员之间操作信息不够流通，后期调整为每日碰头制度后明显改善，建议未来实践从一开始就建立信息共享机制。',
]:
    body(doc, item, first_line=True, space_after=6)

heading(doc, '（三）对课程实践的建议', font_size=14)
for item in [
    '1. 建议在实践前增加1次方法演示课，由指导教师示范切割操作全流程，有助于减少初期因操作失误造成的材料浪费；',
    '2. 建议为各小组提供简易温湿度记录表（纸质或电子），统一记录格式，便于期末汇总横向比较各组数据；',
    '3. 可增设"误差分析"环节，引导同学们对出球率差异进行原因分析，培养科学思维；',
    '4. 建议实践周期结束后组织一次小组成果展示，通过相互交流发现问题、分享经验，提升整体学习效果。',
]:
    body(doc, item, first_line=True, space_after=6)

# ──────────────────────────────────────────────
# 保存
# ──────────────────────────────────────────────
out_path = '/home/user/1/朱顶红切块繁殖与栽培管理实践报告.docx'
doc.save(out_path)
print(f'报告已生成：{out_path}')
