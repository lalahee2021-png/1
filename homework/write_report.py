"""生成课程设计 Word 报告"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import copy, os

doc = Document()

# 全局字体设置 (中文黑体/宋体)
def set_font(run, size=12, bold=False, color=None):
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.name  = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.name = '黑体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        run.font.size  = Pt(16 - 2*level)
    return p

def add_para(doc, text, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0.74 * indent)
    run = p.add_run(text)
    set_font(run, size=11)
    return p

# ─────────────────────────────────────────────────
# 封面
# ─────────────────────────────────────────────────
doc.add_paragraph()
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run('《机车车辆及牵引计算》')
set_font(r, 20, bold=True)

t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run('课程设计报告')
set_font(r, 18, bold=True)

doc.add_paragraph()
for label, val in [('机型', 'HXD1D'), ('货车类型', '重载货车（25t轴重）'),
                   ('学号后两位', '07'), ('编组', '42+7 = 49辆')]:
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(f'{label}：{val}')
    set_font(r, 13)

doc.add_page_break()

# ─────────────────────────────────────────────────
# 0. 已知参数
# ─────────────────────────────────────────────────
add_heading(doc, '一、已知参数', 1)
params_text = """
机车 HXD1D：
  · 轴式 C0-C0，轴重 21t，计算重量 = 黏着重量 = 126 t
  · 最大起动牵引力 420.0 kN，持续牵引力 324.0 kN（80 km/h）
  · 最高速度 160 km/h，全长 22.5 m
  · 单位基本阻力：w₀' = 1.48 + 0.0018v + 0.000304v²  (N/kN)
  · 机车单位起动比阻力：5.0 N/kN
  · 高摩合成闸瓦，紧急制动换算闸瓦压力 340 kN

货车（重载，25t轴重）：
  · 换长 1.3，单轴重 25t，4轴，单车全重 A = 100 t
  · 编组：42 + 07 = 49 辆；列车全重 = 126 + 49×100 = 5026 t
  · 列车全长 = 22.5 + 49×(1.3×13.92) ≈ 909.2 m
  · 单位基本阻力：w₀ = 0.92 + 0.0048v + 0.000125v²  (N/kN，适用至100 km/h)
  · 货车单位起动比阻力：3.5 N/kN
  · 高摩合成闸瓦，紧急制动换算闸瓦压力 195 kN/辆

制动参数：
  · 换算摩擦系数（高摩合成闸瓦）：φ(v) = 0.35×(100+v)/(100+3v)
  · 常用制动力系数：λ = 0.83（正常）；0.5（进站）
  · 基础制动传动效率：机车 0.85，货车 0.90

线路（A→B，由EMF坡道图提取）：
  第1段: 1000m, 0‰（A站出站平坡，限速45km/h）
  第2段: 1500m, 2‰
  第3段:  600m, 6‰
  第4段:  400m, 3‰
  第5段: 1200m, 10‰（限速60km/h）
  第6段:  700m, 7‰
  第7段: 2240m, 15‰（=2100+07×20，最陡段）
  第8段: 1300m, 7‰
  第9段: 1000m, 0‰（B站进站平坡，限速45km/h）
  全程合计: 9940 m
""".strip()

for line in params_text.split('\n'):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(line)
    set_font(r, 10.5)

# ─────────────────────────────────────────────────
# ①
# ─────────────────────────────────────────────────
add_heading(doc, '二、问题①  每辆货车满载时自重与载重之和', 1)
add_para(doc, '重载货车为25t轴重4轴货车，满载时每辆车的自重与载重之和即为整车总重：', 1)
add_para(doc, 'A = 轴重 × 轴数 = 25 t × 4 = 100 t', 2)
add_para(doc, '因此 A = 100 t。', 1)
add_para(doc, '全列车：机车126t + 49辆×100t = 5026 t；列车全长约909.2 m。', 1)

# ─────────────────────────────────────────────────
# ②
# ─────────────────────────────────────────────────
add_heading(doc, '三、问题②  各工况合力曲线', 1)
add_para(doc, '列车编组：HXD1D + 49辆100t重载货车，在平直道（i=0）上各工况合力（净力）'
         '随速度的变化曲线如下图所示。', 1)
add_para(doc, '计算公式：', 1)
add_para(doc, '  列车总基本阻力：W(v) = [126·w₀\'(v) + 4900·w₀(v)] / 1000  (kN)', 2)
add_para(doc, '  牵引净力：F_net = F(v) − W(v)', 2)
add_para(doc, '  惰行净力：F_net = −W(v)', 2)
add_para(doc, '  电制动净力：F_net = −Fe(v) − W(v)', 2)
add_para(doc, '  常用制动净力：F_net = −λ·Fb(v) − W(v)   (λ=0.83)', 2)

if os.path.exists('fig2_force_curves.png'):
    doc.add_picture('fig2_force_curves.png', width=Inches(5.8))
    cap = doc.add_paragraph('图1  列车各工况合力曲线（平直道，49辆货车）')
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(cap.runs[0], 10)

# ─────────────────────────────────────────────────
# ③
# ─────────────────────────────────────────────────
add_heading(doc, '四、问题③  速度-时间曲线与区间运行时分', 1)
add_para(doc, '采用数值仿真（步长0.5s）模拟全程运行。控制策略：', 1)
add_para(doc, '· 牵引阶段（T）：全力牵引，速度不超过允许限速；', 2)
add_para(doc, '· 惰行阶段（C）：不施力，依靠行驶阻力减速；', 2)
add_para(doc, '· 电制动阶段（E）：利用前瞻距离在限速点前减速；', 2)
add_para(doc, '· 混合制动阶段（B）：常用制动（λ=0.83），近站用λ=0.5减速至停车。', 2)
add_para(doc, '限速考虑列车长度：限速区有效范围从区段起点延伸至区段终点+909m（列车长），'
         '前端进入限速区立即生效，尾部完全离开后解除。', 1)

if os.path.exists('fig3_vt_curve.png'):
    doc.add_picture('fig3_vt_curve.png', width=Inches(5.8))
    cap = doc.add_paragraph('图2  速度-时间(上)及速度-距离(下)曲线，红=牵引，绿=惰行，蓝=电制动，紫=混合制动')
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(cap.runs[0], 10)

add_para(doc, '计算结果：区间运行时分 ≈ 746 s（约 12.4 分钟）。', 1)

# ─────────────────────────────────────────────────
# ④
# ─────────────────────────────────────────────────
add_heading(doc, '五、问题④  最限制坡道能否启动', 1)
add_para(doc, '最限制坡道为路线中最陡坡段：15‰（第7段，长2240m）。', 1)
add_para(doc, '起动阻力：', 1)
add_para(doc, '  列车起动基本阻力 = (126×5.0 + 4900×3.5) / 1000 = (630+17150)/1000 = 17.78 kN', 2)
add_para(doc, '  坡道阻力 = 5026×15/1000 = 75.39 kN', 2)
add_para(doc, '  合计起动阻力 = 17.78 + 75.39 = 93.17 kN', 2)
add_para(doc, '机车最大起动牵引力（规格值，已含黏着限制）= 420.0 kN', 1)
add_para(doc, '  420.0 kN >> 93.17 kN，富余牵引力 = 326.8 kN', 2)
add_para(doc, '结论：列车在最限制坡道（15‰）停车后，能够顺利启动。✓', 1)

add_para(doc, '黏着验算：', 1)
add_para(doc, '  所需黏着系数 μ = 420/(126×9.81) ≈ 0.340', 2)
add_para(doc, '  干燥轨面典型黏着系数为0.25~0.40，0.340在正常范围内，不会发生空转。', 2)

# ─────────────────────────────────────────────────
# ⑤
# ─────────────────────────────────────────────────
add_heading(doc, '六、问题⑤  关门车紧急制动距离检算', 1)
add_para(doc, '条件：3辆关门车（制动失效），最高运行速度100km/h，'
         '技规要求紧急制动距离≤1100m。', 1)
add_para(doc, '计算方法（换算法）：', 1)
add_para(doc, '  紧急制动力 = 电制动力Fe(v) + 机械制动力Fb(v, 关门=3辆)', 2)
add_para(doc, '  Fb = [340 + 46×195] × φ(v)  （机车+46辆有效货车）', 2)
add_para(doc, '  φ(v) = 0.35×(100+v)/(100+3v)   高摩合成闸瓦换算摩擦系数', 2)
add_para(doc, '  总制动减速力 = Fe+Fb + 列车基本阻力 + 坡道阻力', 2)
add_para(doc, '  制动距离通过数值积分(步长0.1s)求解。', 2)

# 结果表格
table = doc.add_table(rows=1, cols=5)
table.style = 'Table Grid'
hdr = table.rows[0].cells
for i, txt in enumerate(['坡度(‰)', '坡长(m)', '制动距离(m)', '是否超限', '建议限速(km/h)']):
    hdr[i].text = txt
    for para in hdr[i].paragraphs:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in para.runs:
            run.font.bold = True
            set_font(run, 10)

results = [
    (0,    1000, 941.3,  False, None),
    (2,    1500, 936.7,  False, None),
    (6,     600, 927.6,  False, None),
    (3,     400, 934.4,  False, None),
    (10,   1200, 918.7,  False, None),
    (7,     700, 925.3,  False, None),
    (15,   2240, 907.8,  False, None),
    (7,    1300, 925.3,  False, None),
    (0,    1000, 941.3,  False, None),
]

for grade, length, dist, exceed, vlim in results:
    row = table.add_row().cells
    vals = [str(grade), str(length), f'{dist:.1f}',
            '超限 ✗' if exceed else '合格 ✓',
            f'{vlim:.1f}' if vlim else '不限速']
    for i, v in enumerate(vals):
        row[i].text = v
        for para in row[i].paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                set_font(run, 10)
                if exceed and i == 3:
                    run.font.color.rgb = RGBColor(255, 0, 0)

cap = doc.add_paragraph('表1  各坡道紧急制动距离检算结果（v=100km/h，3辆关门车）')
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(cap.runs[0], 10)

if os.path.exists('fig5_brake_check.png'):
    doc.add_picture('fig5_brake_check.png', width=Inches(5.5))
    cap = doc.add_paragraph('图3  各坡道紧急制动距离检算（红线=1100m限制）')
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(cap.runs[0], 10)

add_para(doc, '结论：在全部坡道上，含3辆关门车时的紧急制动距离均小于1100m，'
         '不需要对任何坡段降低最高运行速度。✓', 1)

# ─────────────────────────────────────────────────
# 附录：源代码
# ─────────────────────────────────────────────────
doc.add_page_break()
add_heading(doc, '附录：源代码（Python）', 1)
add_para(doc, '编程语言：Python 3  依赖库：numpy, matplotlib, scipy', 1)

with open('traction_calc.py', 'r', encoding='utf-8') as f:
    code = f.read()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.2)
r = p.add_run(code)
r.font.name    = 'Courier New'
r.font.size    = Pt(7.5)
r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Courier New')

doc.save('课程设计报告_学号07.docx')
print("报告已保存: 课程设计报告_学号07.docx")
