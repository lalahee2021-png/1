from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
section = doc.sections[0]
section.top_margin    = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin   = Cm(3.17)
section.right_margin  = Cm(3.17)

# ── Helpers ───────────────────────────────────────────────────────────────────
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

def add_table(doc, headers, rows_data, col_widths, header_align_left_cols=()):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, (cell, h, w) in enumerate(zip(hdr, headers, col_widths)):
        cell.width = w
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run(run, size=11, bold=True)
    for row in rows_data:
        cells = table.add_row().cells
        for i, (cell, val, w) in enumerate(zip(cells, row, col_widths)):
            cell.width = w
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i in header_align_left_cols else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            set_run(run, size=10.5)
    return table

# ══════════════════════════════════════════════════════════════════════════════
# Title
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '福建省漳州市9～12月蔬菜无土栽培方案设计', size=16, space_before=0)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after  = Pt(6)
r1 = p.add_run('方案简介：')
set_run(r1, size=12, bold=True)
r2 = p.add_run(
    '漳州地处福建南部，属南亚热带海洋性季风气候，9月仍受夏季高温余威影响（日均气温可达28～32℃），'
    '10～11月转入秋高气爽期（日均22～26℃，昼夜温差加大，光照充足），12月气温进一步下降'
    '（日均15～18℃，最低多在8～10℃，基本无霜冻）。本方案针对该阶段气候由热转凉、昼夜温差逐渐拉大的特点，'
    '选取果菜类（番茄、辣椒）与叶菜类（生菜、小白菜）共4种蔬菜，采用基质袋（槽）培方式，'
    '从基质配制、育苗容器、种子前处理与催芽、定植后肥水管理、温湿度调控等环节制定无土栽培技术方案，'
    '实现9月降温保苗、10～11月旺盛生长、12月保温防寒的全季管理目标。'
)
set_run(r2, size=12)
p.paragraph_format.first_line_indent = Pt(0)

doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# 一、蔬菜种类选择及气候适应性
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '一、蔬菜种类选择及气候适应性分析', size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

add_para(doc,
    '结合漳州9～12月温光条件及市场需求，选取果菜、叶菜各2种，具体品种及茬口安排见下表：',
    size=12, first_line=True)

headers1 = ['类别', '蔬菜种类', '推荐品种', '生育期（天）', '播种期建议', '气候适应性说明']
col_w1 = [Cm(1.6), Cm(2.0), Cm(3.2), Cm(2.0), Cm(2.4), Cm(4.2)]
rows1 = [
    ('果菜', '番茄', '浙粉702、金鹏8号',
     '90～110', '9月上中旬育苗',
     '喜温不耐高温，9月育苗期需遮阳降温，10～12月低温有利于坐果和糖分积累'),
    ('果菜', '辣椒', '博辣红牛、湘研15号',
     '110～130', '9月上旬育苗',
     '喜温怕涝又怕旱，苗期忌高温多雨引发猝倒病；较番茄更不耐低温，12月需加强保温'),
    ('叶菜', '生菜', '意大利生菜、奶油生菜',
     '35～45', '9月下旬起分期播种',
     '喜冷凉，9月高温易先期抽薹，需低温催芽；10～12月为最适生长期'),
    ('叶菜', '小白菜（上海青）', '四季青、矮脚黄',
     '30～40', '9月起可分期播种',
     '适应性广，耐热又耐寒，发芽出苗快，9～12月均可持续分批种植'),
]
add_table(doc, headers1, rows1, col_w1, header_align_left_cols=(2, 5))
doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# 二、栽培基质选择
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '二、栽培基质选择', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

substrate = [
    ('（一）果菜类基质（番茄、辣椒）',
     '采用椰糠∶珍珠岩∶蛭石 = 5∶3∶2（体积比）混合基质，装入基质袋或栽培槽，袋高18～20 cm。'
     '该配比透气性好、保水保肥能力强，适合根系发达、需水需肥量大的果菜；'
     '辣椒根系较番茄浅、既怕涝又怕旱，可将珍珠岩比例适当提高至4份（椰糠∶珍珠岩∶蛭石=5∶4∶1）以增强排水性，防止沤根。'
     '使用前需暴晒或高温蒸汽消毒2小时以上，并用清水反复淋洗，将EC值调至0.3 mS/cm以下、pH调至5.8～6.5备用。'),
    ('（二）叶菜类基质（生菜、小白菜）',
     '采用椰糠∶蛭石 = 3∶1（体积比）的轻型基质，装入浅槽（槽深10～12 cm）或育苗盘直接定植，'
     '也可采用浮板毛管水培法。叶菜根系浅、生育期短，基质应偏疏松、保水性稍强，'
     '同样需预先淋洗调节pH至5.8～6.2、EC降至0.3 mS/cm以下。'),
    ('（三）基质使用注意事项',
     '基质应每季轮换或消毒后重复利用，重茬使用前用0.1%高锰酸钾溶液或太阳能高温闷棚消毒，'
     '防止土传病害积累；栽培袋（槽）底部应开排液孔，坡度1%～2%便于余液回收。'),
]
for title, body in substrate:
    add_para(doc, title, size=12, bold=True, space_before=4, space_after=2)
    add_para(doc, body, size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 三、播种容器选择
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '三、播种容器选择', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

headers3 = ['蔬菜类别', '育苗容器', '孔径/规格', '选择依据']
col_w3 = [Cm(2.0), Cm(3.0), Cm(3.0), Cm(6.4)]
rows3 = [
    ('番茄', '聚苯乙烯（PS）穴盘', '50孔或72孔穴盘',
     '果菜幼苗根系发达、苗龄较长（25～35天），大孔径穴盘基质容量大，可减少中途换钵，利于培育壮苗'),
    ('辣椒', '聚苯乙烯（PS）穴盘', '72孔穴盘',
     '辣椒发芽慢、苗期比番茄更长（约50～60天），根系纤细怕伤根，宜用72孔穴盘一次成苗、减少移栽伤根'),
    ('生菜、小白菜', '聚苯乙烯（PS）穴盘', '128孔或200孔穴盘',
     '叶菜幼苗小、苗龄短（15～20天），小孔径穴盘可密植育苗、节省苗床面积，出苗整齐便于分批定植'),
]
add_table(doc, headers3, rows3, col_w3, header_align_left_cols=(3,))
doc.add_paragraph()

add_para(doc,
    '穴盘每次使用前均需用0.1%高锰酸钾溶液或84消毒液（1∶200稀释）浸泡消毒15～20分钟，'
    '清水冲净后晾干备用；穴孔填充基质应压实至八分满，播种后覆盖一层薄基质（厚约0.5～1 cm）。',
    size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 四、种子前处理
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '四、种子前处理', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

add_para(doc,
    '播种前先行选种：剔除秕籽、破损籽及杂质，可用3%～5%食盐水漂选，去除漂浮的不实种子，清水冲洗后晾干，'
    '再按不同蔬菜种类分别进行消毒处理：',
    size=12, first_line=True, space_before=0, space_after=4)

headers4 = ['蔬菜种类', '消毒方法', '处理要点']
col_w4 = [Cm(2.4), Cm(4.6), Cm(7.4)]
rows4 = [
    ('番茄', '55℃温汤浸种15分钟，或10%磷酸三钠溶液浸种20分钟',
     '温汤浸种时不断搅拌水温至30℃以下，可有效钝化种传病毒及早疫病病原；处理后清水冲净并继续浸种4～6小时'),
    ('辣椒', '55℃温汤浸种15～20分钟，或1%硫酸铜溶液浸种5分钟，或10%磷酸三钠溶液浸种10分钟',
     '温汤浸种时不断搅拌至水温降至30℃，可减轻病毒病、炭疽病等种传病害；'
     '消毒后清水冲净，再用清水浸种4～6小时使种子充分吸胀'),
    ('生菜', '常温清水浸种2～3小时，或0.3%碳酸氢钠（小苏打）溶液浸种20分钟',
     '生菜种皮薄嫩，忌用高温烫种，消毒宜温和，浸种后捞出沥干进入催芽环节'),
    ('小白菜', '55℃温汤浸种15分钟，或50%多菌灵可湿性粉剂500倍液浸种30分钟',
     '十字花科种子较耐处理，消毒后清水冲洗干净，沥干水分后即可催芽'),
]
add_table(doc, headers4, rows4, col_w4, header_align_left_cols=(1, 2))
doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# 五、催芽处理
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '五、催芽处理', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

sprout = [
    ('（一）番茄', '浸种消毒后用湿纱布包裹，置于25～30℃恒温催芽箱（或用泡沫箱+温水袋保温）中催芽，'
     '每天用清水淘洗种子1～2次防止发酵酸败，约2～3天种子露白即可播种。'),
    ('（二）辣椒', '浸种消毒后用湿纱布包裹，置于28～30℃条件下催芽，辣椒种皮较厚、发芽较番茄缓慢，'
     '需每天用温水淘洗种子1～2次并保持湿润透气，约5～7天种子露白（比番茄多2～4天），'
     '发芽不整齐时可适当延长1～2天，露白后及时播种，避免胚根过长受损。'),
    ('（三）生菜（应对9月高温的关键环节）', '生菜种子具有高温休眠特性，气温超过25℃时发芽率显著下降。'
     '9月漳州气温偏高，浸种后应将种子置于4～10℃冰箱冷藏室或阴凉处低温处理24小时打破休眠，'
     '再转入15～20℃条件下变温催芽2～3天，种子露白后及时播种；10月以后气温回落，可直接常温催芽。'),
    ('（四）小白菜', '十字花科种子发芽不苛求低温，浸种后置于20～25℃条件下催芽，'
     '约1～2天即可露白出芽，可直接播种，无需特殊变温处理。'),
]
for title, body in sprout:
    add_para(doc, title, size=12, bold=True, space_before=4, space_after=2)
    add_para(doc, body, size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 六、肥水管理
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '六、肥水管理', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

add_para(doc, '（一）营养液配方（以山崎配方为基础调整）', size=12, bold=True, space_before=4, space_after=2)

headers6 = ['配方（每1000 L母液用量）', '硝酸钙', '硝酸钾', '磷酸二氢钾', '硫酸镁', 'EDTA铁及微量元素']
col_w6 = [Cm(3.6), Cm(2.0), Cm(2.0), Cm(2.2), Cm(2.0), Cm(2.2)]
rows6 = [
    ('番茄、辣椒（果菜通用液）', '950 g', '810 g', '340 g', '500 g', '按说明书剂量加入'),
    ('生菜、小白菜（叶菜通用液）', '590 g', '505 g', '170 g', '370 g', '按说明书剂量加入'),
]
add_table(doc, headers6, rows6, col_w6)
doc.add_paragraph()

add_para(doc,
    '辣椒对氮、磷需求与番茄相近，可共用同一果菜通用液；但辣椒挂果期需钾量较高、耐盐性较番茄弱，'
    '实际使用时可在通用液基础上适当增补硫酸钾、并较番茄降低10%左右施用浓度，以促进果实膨大着色、防止烧根。',
    size=12, first_line=True, space_before=0, space_after=4)

add_para(doc, '（二）EC、pH 及灌溉制度', size=12, bold=True, space_before=4, space_after=2)

headers6b = ['蔬菜类别', '生长阶段', 'EC值（mS/cm）', 'pH', '灌溉频率与要点']
col_w6b = [Cm(2.2), Cm(2.2), Cm(2.2), Cm(1.4), Cm(5.6)]
rows6b = [
    ('番茄', '幼苗期', '1.2～1.5', '6.0～6.5', '晴天1～2次/日，基质见干见湿，防止苗期徒长'),
    ('番茄', '开花坐果期', '2.0～2.5', '6.0～6.5', '晴天2～3次/日，每次滴灌至回液率10%～20%'),
    ('番茄', '结果盛期', '2.5～3.0', '6.0～6.5', '晴天3～4次/日，阴雨天减为1次，冬季（12月）适当降低灌溉量'),
    ('辣椒', '幼苗期', '1.0～1.3', '6.0～6.5', '晴天1～2次/日，基质保持湿润偏干，避免徒长和沤根'),
    ('辣椒', '开花坐果期', '1.8～2.2', '6.0～6.5', '晴天2次/日，忌大水大肥，防止落花落果'),
    ('辣椒', '结果盛期', '2.0～2.5', '6.0～6.5', '晴天2～3次/日，辣椒耐盐性弱于番茄，EC不宜超过2.5，冬季（12月）适当减量'),
    ('生菜、小白菜', '全生育期', '1.2～1.8', '5.8～6.2', '晴天1～2次/日，需肥量小，忌积水，阴雨天可停灌1天'),
]
add_table(doc, headers6b, rows6b, col_w6b, header_align_left_cols=(4,))
doc.add_paragraph()

add_para(doc,
    '灌溉均采用滴灌或微喷方式定时定量供给营养液，每日结合天气及基质湿度灵活调整；'
    '每隔7～10天检测一次营养液EC、pH，及时补充或更换，防止盐分累积或养分失衡。',
    size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 七、温度湿度管理
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '七、温度湿度管理', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

headers7 = ['时段', '漳州气候特点', '温度管理措施', '湿度管理措施']
col_w7 = [Cm(1.4), Cm(3.2), Cm(4.4), Cm(4.6)]
rows7 = [
    ('9月', '日均气温28～32℃，仍处夏季高温余威期，午后易超35℃',
     '大棚覆盖遮光率30%～50%遮阳网，加强顶部与侧面通风，中午前后适当揭膜降温；'
     '辣椒育苗期忌高温高湿，需防猝倒病，遮阳的同时加强通风；育苗期生菜种子须冷藏低温处理打破休眠',
     '空气湿度控制在80%～90%利于出苗，午后可对棚内地面喷雾降温增湿，但避免叶面积水'),
    ('10～11月', '日均气温22～26℃，昼夜温差加大，光照充足，气候最适宜',
     '基本无需人工增温降温，白天正常通风换气，夜间视气温酌情闭棚保温',
     '成株期果菜适宜湿度60%～70%，叶菜65%～75%，注意通风排湿，减少病害发生'),
    ('12月', '日均气温15～18℃，早晚最低可至8～10℃，基本无霜冻但夜间偏凉',
     '番茄夜间加盖二层或三层大棚膜保温，必要时使用暖风机或电热加温设备；'
     '辣椒比番茄更不耐低温，生长适温低于15℃即明显减缓、低于10℃易受寒害，需在番茄保温措施基础上'
     '再加盖一层棚膜或增设小拱棚，确保夜间棚温不低于12℃；晴天上午揭膜通风排湿，防止棚温骤升；'
     '叶菜耐寒性较强，一般覆盖一层薄膜即可',
     '棚内湿度控制在60%～70%，避免夜间闷湿诱发灰霉病、霜霉病；辣椒忌高湿环境，'
     '湿度过大易诱发疫病、炭疽病，应加强通风排湿，早晨适当延迟揭膜以防冷凝水滴落伤苗'),
]
add_table(doc, headers7, rows7, col_w7, header_align_left_cols=(1, 2, 3))
doc.add_paragraph()

add_para(doc,
    '全程应加强巡查，发现基质过湿或叶面结露及时通风排湿；高温期结合遮阳与喷雾降温，'
    '低温期以覆盖保温为主、加温为辅，避免昼夜温差过大导致果菜裂果或叶菜抽薹；'
    '辣椒对低温和高湿均较敏感，12月应重点加强保温与通风排湿的平衡，防止落花落果及病害发生。',
    size=12, first_line=True, space_before=0, space_after=4)

# ══════════════════════════════════════════════════════════════════════════════
# 八、采收安排
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '八、采收安排', size=13, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

add_para(doc,
    '生菜、小白菜播种后30～45天陆续采收，可自9月下旬至12月分批分期播种、滚动采收，'
    '实现叶菜周年不间断供应；番茄9月上中旬育苗，10月中下旬定植，'
    '于12月上中旬陆续始收；辣椒因苗期较长，9月上旬即需育苗，10月下旬定植，'
    '于12月中下旬陆续始收，两种果菜采收期均可持续至次年1～2月，跨越整个秋冬季生产周期。',
    size=12, first_line=True, space_before=0, space_after=4)

# ── Save ──────────────────────────────────────────────────────────────────────
out_path = '/home/user/1/福建省漳州市9-12月蔬菜无土栽培方案设计.docx'
doc.save(out_path)
print(f'Saved: {out_path}')
