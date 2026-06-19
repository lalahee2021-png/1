from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── 页面设置 ──────────────────────────────────────────────
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.left_margin   = Cm(3.18)
section.right_margin  = Cm(3.18)
section.top_margin    = Cm(2.54)
section.bottom_margin = Cm(2.54)

# ── 默认正文字体（全局）───────────────────────────────────
style_normal = doc.styles['Normal']
style_normal.font.name = '宋体'
style_normal.font.size = Pt(12)
style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def set_font(run, name='宋体', size=12, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_paragraph(text='', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0,
                  name='宋体', size=12, bold=False, color=None, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * indent)
    if text:
        run = p.add_run(text)
        set_font(run, name, size, bold, color)
    return p

def add_heading(text, level=1):
    """自定义标题，不用内置 Heading 样式避免编号干扰"""
    sizes  = {1: 16, 2: 14, 3: 13, 4: 12}
    bolds  = {1: True, 2: True, 3: True, 4: True}
    before = {1: 18, 2: 12, 3: 10, 4: 8}
    after  = {1: 6,  2: 6,  3: 4,  4: 4}
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(before[level])
    p.paragraph_format.space_after  = Pt(after[level])
    run = p.add_run(text)
    set_font(run, '黑体', sizes[level], bolds[level])
    return p

def add_table(headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_font(run, '黑体', 11, True)
        cell._tc.get_or_add_tcPr()

    # 数据行
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            set_font(run, '宋体', 11)

    # 列宽
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)

    # 底部间距
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table

def add_flow_box(text):
    """流程图用灰色背景段落模拟"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    set_font(run, '仿宋', 11)
    # 加底纹
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)
    return p

# ════════════════════════════════════════════════════════
#  封  面
# ════════════════════════════════════════════════════════
for _ in range(4):
    doc.add_paragraph()

add_paragraph('朱顶红切块繁殖与栽培管理技术实践报告',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='黑体', size=18, bold=True,
              space_before=0, space_after=40)

for label, val in [('小组名称', '第X组'),
                   ('小组成员', 'XXX、XXX、XXX'),
                   ('实践时间', '2026年3月10日 — 2026年6月10日'),
                   ('实践地点', '植物园温室大棚')]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r1 = p.add_run(f'{label}：')
    set_font(r1, '黑体', 13, True)
    r2 = p.add_run(val)
    set_font(r2, '宋体', 13)

doc.add_page_break()

# ════════════════════════════════════════════════════════
#  一、实践目的
# ════════════════════════════════════════════════════════
add_heading('一、实践目的', 1)

add_paragraph(
    '朱顶红（Hippeastrum vittatum）属石蒜科多年生草本植物，以其硕大艳丽的花朵和较强的适应性深受园艺爱好者喜爱。'
    '鳞茎切块繁殖是利用其鳞茎内部储存的丰富养分及潜在芽点，通过人工切割手段刺激每一切块独立发育成新植株的一种无性繁殖方式，'
    '具有操作简便、繁殖系数高、遗传性状稳定等优点，在规模化育苗生产中具有重要应用价值。',
    indent=2, space_after=6)

add_paragraph('本次实践旨在达成以下目标：', indent=2, space_after=4)

goals = [
    '通过亲手操作，熟练掌握朱顶红鳞茎切块繁殖的全流程技术，包括种球筛选、消毒切割、伤口处理及扦插养护；',
    '深入理解朱顶红各生长阶段的生理需求，学会依据温度、湿度、光照等环境因素进行针对性调控；',
    '系统学习朱顶红日常栽培养护方法，掌握常见病虫害的识别特征与综合防治策略；',
    '拓宽视野，了解朱顶红除切块法之外的其他繁殖途径（如分球繁殖、播种繁殖、组织培养等），为后续专业学习奠定基础。',
]
for i, g in enumerate(goals, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f'{i}. {g}')
    set_font(run, '宋体', 12)

add_paragraph(
    '实践方案概述：本次实践以小组为单位，在植物园温室大棚内选取成熟朱顶红母球，分别采用十字切割法和楔形切割法进行切块扦插，'
    '同步开展盆栽管理实验，观察不同水肥处理对鳞茎生长发育的影响，最终汇总数据撰写报告。',
    indent=2, space_before=6, space_after=6)

# ════════════════════════════════════════════════════════
#  二、实践内容
# ════════════════════════════════════════════════════════
add_heading('二、实践内容', 1)
add_heading('（一）前期准备', 2)
add_heading('1. 材料与器具清点', 3)

add_paragraph('实践开始前，小组成员共同核查所需材料（见表1）。', indent=2, space_after=4)

add_table(
    ['类别', '名称', '规格/说明'],
    [
        ['试验材料', '朱顶红成熟母球', '直径8~10 cm，无病斑，鳞茎盘完整'],
        ['试验材料', '朱顶红子球',   '直径3~4 cm，用于对比观察'],
        ['切割工具', '锋利菜刀/枝剪', '使用前用75%酒精擦拭消毒'],
        ['消毒药品', '75%酒精',       '用于工具及种球表面消毒'],
        ['消毒药品', '0.1%多菌灵溶液', '切块浸泡消毒'],
        ['消毒药品', '草木灰/硫磺粉', '切口涂抹，防止继发感染'],
        ['扦插基质', '珍珠岩：河沙：泥炭=1:1:4', '消毒处理后使用'],
        ['盆栽基质', '腐叶土：园土：河沙=5:3:2', '加入腐熟有机肥'],
        ['容器',     '育苗盘、20 cm花盆', '—'],
        ['其他',     '保鲜膜、喷壶、标签牌', '—'],
    ],
    col_widths=[2.5, 4, 7.5]
)
add_paragraph('表1 实践材料清单', align=WD_ALIGN_PARAGRAPH.CENTER,
              name='黑体', size=11, space_before=0, space_after=8)

add_heading('2. 种球初步处理', 3)
add_paragraph(
    '领取母球后，逐一检查外观：剥除已干枯发褐的外层鳞片，露出乳白色、质地紧实的健康鳞片；仔细观察鳞茎盘是否完整，'
    '基部是否存在腐烂迹象；同时确认顶芽是否饱满充实。不符合要求的种球予以剔除，不参与切块实验。',
    indent=2, space_after=6)
add_paragraph(
    '入选种球提前1周停止浇水，使其充分干燥，以降低切割时汁液外流造成的腐烂风险，并使鳞片组织更加致密，便于清晰切割。',
    indent=2, space_after=6)

add_paragraph('【图1：种球外观检查（拍摄日期：2026-03-10）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

# ── 切块繁殖操作 ──────────────────────────────────────
add_heading('（二）切块繁殖操作', 2)
add_heading('1. 切割时机把握', 3)
add_paragraph(
    '本次实践于3月中旬启动，此时气温逐渐回暖，鳞茎经历冬季休眠后开始由内向外积累萌发动力，'
    '切块后愈伤组织形成速度快、萌芽率高，是最佳操作窗口期。',
    indent=2, space_after=6)

add_heading('2. 切割方法实施', 3)
add_paragraph('消毒完毕后，在操作台上铺好消毒纸巾，将种球竖直放置，切割步骤如下：', indent=2, space_after=4)

add_heading('（1）预切根盘', 4)
add_paragraph(
    '用消过毒的菜刀沿鳞茎基部切去老化干枯的根须，保留完整的鳞茎盘组织（即白色圆盘状部分），'
    '此为后续形成新根的关键结构。',
    indent=2, space_after=6)

add_heading('（2）纵向主切', 4)
add_paragraph(
    '以鳞茎顶芽为中心，由上至下纵向切开，将鳞茎平均分成4等份（十字切割法）；对直径较大的母球，'
    '可进一步将每份再纵切，获得8~16个楔形切块（楔形切割法）。',
    indent=2, space_after=4)
add_paragraph(
    '关键要领：每刀下去须一气呵成，切面要平整光滑，避免来回锯拉撕裂鳞片纤维。每一切块下端必须保留约0.5 cm宽的鳞茎盘组织，'
    '否则该切块将因无法生根而彻底报废。',
    indent=2, bold=False, space_after=6)
add_paragraph('【图2：十字切割完成后的切块（拍摄日期：2026-03-12）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

add_heading('（3）伤口消毒与晾干', 4)
add_paragraph(
    '切块完成后立即投入0.1%多菌灵溶液中浸泡约10分钟，取出后平铺于报纸上，置于通风阴凉处晾置约1.5天，'
    '待切口表面呈现轻微干缩、不再湿润时，在切口断面均匀涂抹草木灰粉，形成物理隔离保护层，有效降低霉菌侵入风险。',
    indent=2, space_after=6)
add_paragraph('【图3：切块晾干及涂灰处理（拍摄日期：2026-03-13）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

add_heading('3. 扦插上床', 3)
add_paragraph(
    '将珍珠岩：河沙：泥炭（1:1:4）的混合基质用0.1%高锰酸钾溶液浸透消毒，晾至湿润不滴水状态后装入育苗盘。'
    '将处理好的切块芽眼朝上平放于基质表面，轻压使切块底部与基质充分接触，切块间距保持约8 cm，避免相互接触传染。'
    '随后用保鲜膜覆盖育苗盘四周，保留小缝隙换气，整盘移入温室黑暗角落（催芽阶段须遮光）。',
    indent=2, space_after=6)
add_paragraph('【图4：切块扦插排布（拍摄日期：2026-03-14）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

add_heading('4. 扦插后养护管理', 3)

items = [
    ('温度', '温室内温度维持在26~28℃之间，每日早晚各记录一次温度数据，发现偏低时及时关闭通风口。'),
    ('湿度', '每日检查基质表面干湿状态，用喷壶在保鲜膜上喷水，维持环境相对湿度在70~80%。禁止直接大量浇水，防止积水导致切块腐烂。'),
    ('光照', '催芽阶段全程遮光，避免光照抑制小种球分化。待切块表面出现白色芽点（约30天后）后，可逐步引入散射光照。'),
    ('施肥', '扦插后第5周，开始以喷雾方式施用稀薄叶面肥（约1/2推荐浓度），每10天一次，促进小鳞茎进一步发育充实。'),
]
for k, v in items:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run(f'{k}：')
    set_font(r1, '黑体', 12, True)
    r2 = p.add_run(v)
    set_font(r2, '宋体', 12)

add_heading('5. 生长观察记录', 3)
add_paragraph('表2 切块繁殖过程观察记录表', align=WD_ALIGN_PARAGRAPH.CENTER,
              name='黑体', size=11, space_before=4, space_after=4)
add_table(
    ['记录日期', '天数', '切块外观变化', '芽点萌发情况', '腐烂情况', '备注'],
    [
        ['2026-03-14', '第1天',  '切口干燥，草木灰附着',   '无',   '0块', '扦插当日'],
        ['2026-03-21', '第7天',  '切口愈伤组织开始形成',   '无',   '1块（伤口未干充分）', '—'],
        ['2026-03-28', '第14天', '部分切块底部出现白色突起','少量', '1块', '腐烂块已移除'],
        ['2026-04-14', '第30天', '多数切块表面可见绿色芽点','明显', '0块', '首次施叶面肥'],
        ['2026-05-14', '第60天', '小鳞茎直径约0.4~0.8 cm', '生长旺盛','0块','—'],
        ['2026-06-10', '第87天', '小鳞茎直径达0.6~1.2 cm', '部分已长出2~3条根','0块','达到移栽标准'],
    ],
    col_widths=[2.4, 1.6, 3.6, 2.8, 2.0, 1.6]
)

add_paragraph('【图5：第30天芽点萌发（拍摄日期：2026-04-14）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=4)
add_paragraph('【图6：第60天小鳞茎生长（拍摄日期：2026-05-14）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

add_heading('繁殖数据统计（本组）', 4)
stats = [
    '参与切块繁殖的母球数：3个',
    '总切块数：32块（十字切割法12块，楔形切割法20块）',
    '成功出球数：28块',
    '腐烂/失败数：4块（其中2块扦插初期腐烂，2块未萌发）',
    '出球率：28 ÷ 32 × 100% = 87.5%',
    '繁殖系数：28 ÷ 3 ≈ 9.3',
    '腐烂率：2 ÷ 32 × 100% = 6.25%',
]
for s in stats:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(f'· {s}')
    set_font(run, '宋体', 12)
    if '%' in s or '系数' in s:
        run.font.bold = True

# ── 盆栽管理 ──────────────────────────────────────────
add_heading('（三）盆栽栽培管理实验', 2)
add_heading('1. 种植准备', 3)
add_paragraph(
    '将腐叶土、园土、河沙按5:3:2体积比充分混合，加入适量腐熟有机肥，装入20 cm口径花盆，轻压基质使之平整。',
    indent=2, space_after=6)

add_heading('2. 种球定植', 3)
add_paragraph(
    '选取直径约4 cm的子球，以鳞茎顶部露出土面约1/3为准确定种植深度，避免深埋导致顶芽出土困难，'
    '也避免过浅导致根部不稳。种植后浇透水，直至盆底排水孔渗出水为止。',
    indent=2, space_after=6)
add_paragraph('【图7：朱顶红子球定植（拍摄日期：2026-03-15）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

add_heading('3. 日常养护记录', 3)
care = [
    ('光照管理',
     '将花盆置于温室内光照充足位置，春秋季每日可接受直射光6小时以上；进入4月下旬后温室内温度升高，'
     '在午间11:00~14:00期间辅以遮阳网，防止强光灼伤叶片。'),
    ('温度记录',
     '生长期间温室温度日均维持在20~24℃，符合朱顶红适宜生长温度范围（18~25℃）。'),
    ('水分管理',
     '遵循"见干见湿"原则，用手指触探土表下2 cm处，感觉偏干时方可浇水。平均每周浇水1~2次，'
     '梅雨季节适当减少浇水量，防止根部长期积水缺氧腐烂。'),
    ('施肥记录',
     '种植后第三周起开始施肥，每15天施用一次稀薄饼肥水（稀释比例约1:100）；'
     '花芽分化明显后改施磷钾肥为主的配方肥，以促进花葶粗壮、花色鲜艳。'),
]
for k, v in care:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(5)
    r1 = p.add_run(f'{k}：')
    set_font(r1, '黑体', 12, True)
    r2 = p.add_run(v)
    set_font(r2, '宋体', 12)

add_paragraph('表3 盆栽管理实验数据记录（随机10盆，处理A：正常水肥；处理B：减半水肥）',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='黑体', size=11,
              space_before=8, space_after=4)
add_table(
    ['测量指标', '处理A均值', '处理B均值', '说明'],
    [
        ['鳞茎纵径（cm）',    '7.2', '5.8', '收获期测量'],
        ['鳞茎横径（cm）',    '8.1', '6.5', '—'],
        ['鳞茎盘直径（cm）',  '3.4', '2.9', '—'],
        ['单个鳞茎鲜重（g）', '186', '134', '—'],
        ['叶片长度（cm）',    '42.3','31.7','最长叶测量'],
        ['叶片宽度（cm）',    '4.8', '3.6', '—'],
    ],
    col_widths=[4.5, 3, 3, 3.5]
)
add_paragraph('【图8：处理A与处理B生长对比（拍摄日期：2026-05-20）】',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='仿宋', size=11,
              color=(128,128,128), space_after=10)

add_heading('4. 病虫害观察与防治', 3)
add_paragraph('实践期间共观察到以下病虫害发生情况：', indent=2, space_after=4)

pest_items = [
    ('红斑病（赤斑病）',
     '4月下旬，温室内有2盆植株叶片出现椭圆形红褐色斑点，后期斑点中央颜色加深。初步判断为赤斑病（由朱顶红炭疽菌引起）。'
     '处理措施：立即摘除病叶，用50%多菌灵可湿性粉剂800倍液全株喷施，每7天喷一次，连续3次后病情受到控制。'),
    ('介壳虫',
     '5月初，部分植株叶背和茎部发现少量介壳虫，用棉签蘸75%酒精逐一擦除，并喷施10%吡虫啉1500倍液进行防治，'
     '约2周后虫害基本消除。'),
]
for k, v in pest_items:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(5)
    r1 = p.add_run(f'（1）{k}：' if k == '红斑病（赤斑病）' else f'（2）{k}：')
    set_font(r1, '黑体', 12, True)
    r2 = p.add_run(v)
    set_font(r2, '宋体', 12)

# ════════════════════════════════════════════════════════
#  三、实践结论与分析
# ════════════════════════════════════════════════════════
add_heading('三、实践结论与分析', 1)
add_heading('（一）朱顶红栽培与管理技术总结', 2)
add_paragraph('通过本次实践，朱顶红栽培管理的核心要点可归纳为以下几个方面：', indent=2, space_after=4)

conclusions = [
    ('基质选配是根本',
     '朱顶红根系肉质，需氧性强，最忌积水。栽培基质务必选用排水透气性好的配方，本次实验所用腐叶土+园土+河沙（5:3:2）'
     '配比表现良好，植株根系发达，未出现积水烂根现象。若土壤黏性过重，应适当增大河沙比例。'),
    ('光照充足是保障',
     '实验数据显示，处理A（正常光照）下植株叶片更宽更长，鳞茎也明显更加充实。光照不足会导致叶片细长软弱（徒长），'
     '并影响花芽分化质量。建议生长期每日保证6小时以上光照，夏季高温时段须遮阴，以避免叶片焦尖。'),
    ('水肥合理是关键',
     '处理A与处理B的鳞茎鲜重相差约52 g，差异显著，说明水肥充足对鳞茎膨大有积极作用。但过量浇水会造成土壤长期潮湿，'
     '引发根部病害；施肥过浓则会发生肥害烧根。"宁少勿多、少量多次"是较安全的水肥管理原则。'),
    ('休眠管理不可忽视',
     '朱顶红冬季有明显休眠期，此时需控制浇水（每月1~2次），并将温度降至5~10℃低温环境，以使鳞茎得到充分休整，'
     '为次年开花积累养分。不当的越冬管理（如持续高温高湿）会导致鳞茎提前消耗储备、翌年开花量减少或不开花。'),
    ('病虫害预防为主',
     '病虫害一旦发生，治理成本远高于预防。良好的通风换气、合理的水肥控制、定期的植株检查是最有效的预防手段。'
     '发现问题须第一时间隔离病株，对症施药，避免交叉感染蔓延至健康植株。'),
]
for i, (k, v) in enumerate(conclusions, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(14)
    p.paragraph_format.space_after = Pt(5)
    r1 = p.add_run(f'{i}. {k}：')
    set_font(r1, '黑体', 12, True)
    r2 = p.add_run(v)
    set_font(r2, '宋体', 12)

add_heading('（二）切块繁殖操作方法归纳与技术总结', 2)
add_paragraph('操作流程如下：', indent=2, space_after=4)
add_flow_box(
    '选球 → 种球预处理（停水干燥）→ 消毒 → 切割（保留鳞茎盘）\n'
    '  → 伤口消毒（多菌灵浸泡）→ 晾干涂灰 → 基质准备（消毒）\n'
    '    → 扦插（芽眼朝上）→ 遮光控温（25~30℃，湿度70~80%）\n'
    '      → 定期观察记录 → 待出球后移栽定植'
)

add_paragraph('技术关键点总结（见表4）：', indent=2, space_before=8, space_after=4)
add_paragraph('表4 切块繁殖关键技术要点与常见失误对照表',
              align=WD_ALIGN_PARAGRAPH.CENTER, name='黑体', size=11,
              space_before=0, space_after=4)
add_table(
    ['操作环节', '核心要点', '常见失误'],
    [
        ['种球选择', '鳞茎饱满、无病斑、直径≥8 cm', '使用带病或过小种球'],
        ['切割工具', '刀具锋利且彻底消毒', '钝刀撕裂组织、交叉感染'],
        ['切割方式', '每块必须带鳞茎盘，宽度≥1 cm', '切块无根盘、切块过小'],
        ['伤口处理', '多菌灵浸泡→充分晾干→涂草木灰', '未晾干直接扦插'],
        ['扦插摆放', '芽眼向上平放，不埋入基质', '倒放或深埋'],
        ['催芽环境', '25~30℃，遮光，湿度70~80%', '温度不足/光照过强'],
        ['移栽时机', '小种球直径≥0.5 cm，已长出2~3条根', '过早移栽根系未建立'],
    ],
    col_widths=[2.5, 5.5, 6]
)

add_heading('切割方法比较', 4)
methods = [
    ('十字切割法（4块）',
     '操作简便，切块体积大、养分充足，萌发率高，但繁殖数量有限；适合初学者或对繁殖系数要求不高的情况。'),
    ('楔形切割法（8~16块）',
     '繁殖系数高，适合规模化育苗；但操作难度较大，切块较小，对消毒和养护管理要求更严格，腐烂风险相应提高。'),
]
for k, v in methods:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(5)
    r1 = p.add_run(f'· {k}：')
    set_font(r1, '黑体', 12, True)
    r2 = p.add_run(v)
    set_font(r2, '宋体', 12)

add_paragraph(
    '综合来看，在实际操作中可根据种球大小和育苗数量需求灵活组合使用两种方法，'
    '关键是确保每个切块均携带鳞茎盘且操作全程严格消毒。',
    indent=2, space_before=4, space_after=6)

# ════════════════════════════════════════════════════════
#  四、实践体会
# ════════════════════════════════════════════════════════
add_heading('四、实践体会', 1)
add_heading('（一）个人心得体会', 2)
add_paragraph(
    '此次实践让我对朱顶红从一颗鳞茎到多棵子株的"增殖之旅"有了真实而深刻的体验。书本上的"伤口愈合"、"不定芽萌发"'
    '等概念，在亲手操作并每天记录观察数据之后，变得具体而鲜活。',
    indent=2, space_after=6)
add_paragraph(
    '让我印象最深的是第30天揭开保鲜膜的一刻——原本白净光滑的切块表面，长出了米粒大小的白色芽点，那种满足感是单纯翻阅资料所无法替代的。'
    '与此同时，我也真切体会到"消毒"二字的分量：实验初期有2块切块因晾干不充分就上床扦插，短短一周便在基部出现黑色腐斑，'
    '不得不提前淘汰。这让我明白，农业操作中每一个看似琐碎的细节，都与最终结果息息相关。',
    indent=2, space_after=6)

add_heading('（二）小组不足之处', 2)
shortfalls = [
    '记录不够及时系统：部分阶段（尤其是第2~3周）的观察记录出现空缺，导致数据不连贯，后续分析存在一定局限；',
    '照片记录质量参差不齐：部分早期照片拍摄角度不统一、光线较暗，给后期图文排版造成困难；',
    '对照设计不够完善：本次切块实验未设置不同处理浓度的消毒剂对照组，无法量化分析消毒处理对出球率的影响，实验设计的科学性有待加强；',
    '分工协调有待优化：前期小组内任务分配不够清晰，导致部分浇水、施肥操作出现重复执行或遗漏的情况。',
]
for i, s in enumerate(shortfalls, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f'{i}. {s}')
    set_font(run, '宋体', 12)

add_heading('（三）对课程实践的建议', 2)
suggestions = [
    '建议增加操作演示课时：切割环节技术难度较高，希望老师在课前进行一次完整的示范演示，让同学们有更直观的参照；',
    '建议提供标准记录表格：统一的数据记录格式可帮助各小组更规范地收集数据，也便于课后全班汇总比较；',
    '建议开展组间交流展示：实践结束后若能安排各小组分享经验，尤其是遇到的问题与解决方法，将有助于同学们相互取长补短、共同提高；',
    '建议增设繁殖方式对比实验：若课时允许，可同步开展切块法与分球法的平行对比实验，让同学们通过实际数据直观感受不同繁殖方式的优劣差异。',
]
for i, s in enumerate(suggestions, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(28)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f'{i}. {s}')
    set_font(run, '宋体', 12)

# ── 保存 ──────────────────────────────────────────────
out = '/home/user/1/朱顶红切块繁殖与栽培管理实践报告.docx'
doc.save(out)
print(f'已生成：{out}')
