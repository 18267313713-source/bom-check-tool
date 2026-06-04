import pptx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(0x10, 0x2A, 0x43)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0x24, 0x3B, 0x53)
LIGHT_BG = RGBColor(0xF0, 0xF4, 0xF8)

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_slide_number(slide, num, total):
    txBox = slide.shapes.add_textbox(Inches(12.5), Inches(7.0), Inches(0.6), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"{num}/{total}"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(0x88, 0x99, 0xAA)
    p.alignment = PP_ALIGN.RIGHT

def add_header_bar(slide, title_text, subtitle_text=""):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, Inches(1.2))
    shp.fill.solid()
    shp.fill.fore_color.rgb = NAVY
    shp.line.fill.background()
    
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(10), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title_text
    p1.font.size = Pt(24)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    
    if subtitle_text:
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.size = Pt(14)
        p2.font.color.rgb = RGBColor(0x88, 0xAA, 0xCC)

TOTAL_SLIDES = 10

# ... (Keep existing slides) ...

# Helper for adding the new enhancement slide
def add_enhancement_slide(slide_idx, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_header_bar(slide, "八、新增数据质量与防呆校验", "提升数据规范性与系统稳定性的强化规则。")
    add_slide_number(slide, slide_idx, total)
    
    rules = [
        ("1. 关键代码格式净化", "工程物料、物料号、制造物料等代码字段禁止包含中文、空格及特殊符号。仅允许：字母、数字、横杠（-）。"),
        ("2. 防死循环逻辑 (表三)", "BOM 结构中，父级「工程物料」严禁与子级「组件」相同，防止系统展开死循环。"),
        ("3. 防重复工序 (表四)", "同一个「制造物料」下，不允许存在重复的工序编号，确保工艺路线唯一。"),
        ("4. 核心数值范围", "• 净数量、SPM、生产周期：必须 > 0。\n• 废品率、重量类字段：不能 < 0。\n• 模穴数：必须为正整数。\n• 回料百分比：必须在 0-100 之间。")
    ]
    
    y = 1.5
    for title, content in rules:
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(y), Inches(12.3), Inches(1.1))
        shp.fill.solid()
        shp.fill.fore_color.rgb = LIGHT_BG
        shp.line.color.rgb = NAVY
        
        tb = slide.shapes.add_textbox(Inches(0.7), Inches(y + 0.1), Inches(11.9), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        
        p2 = tf.add_paragraph()
        p2.text = content
        p2.font.size = Pt(13)
        p2.font.color.rgb = ACCENT
        p2.space_before = Pt(4)
        y += 1.2

# Update existing slide generation...
# Note: I need to ensure TOTAL_SLIDES is correct in previous definitions? 
# Actually, I can just define them as variables or hardcoded. 
# Since I'm injecting into a file where TOTAL_SLIDES is defined at top, I'll assume it works 
# but wait, the code above redefined TOTAL_SLIDES = 10 (was 8). 
# Let's just make sure the call to add_enhancement_slide is correct.

# Slide 1: Title
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, NAVY)
tb = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10), Inches(3))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "BOM 自校验工具"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER
p2 = tf.add_paragraph()
p2.text = "全五张表格校验规则、枚举值、联动逻辑与公式详解"
p2.font.size = Pt(20)
p2.font.color.rgb = RGBColor(0x88, 0xAA, 0xCC)
p2.alignment = PP_ALIGN.CENTER
p3 = tf.add_paragraph()
p3.text = "技术校验规范"
p3.font.size = Pt(16)
p3.font.color.rgb = RGBColor(0x88, 0xAA, 0xCC)
p3.alignment = PP_ALIGN.CENTER

# Slide 2: Global Rules
s2 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s2, WHITE)
add_header_bar(s2, "一、全局与跨表通用校验规则")
add_slide_number(s2, 2, TOTAL_SLIDES)

rules_2 = [
    ("全局规则", [
        "所有数字的小数位数不得超过 6 位，超出均报错。",
        "所有表格的「导入类型」仅允许填写枚举值：I, D, U。",
    ]),
    ("跨表存在性校验 (以表一为核心)", [
        "表一的「工程物料」必须全部出现在：表二「物料号」、表三「工程物料」、表四「制造物料」、表五「物料」。",
        "若表一在表二、三、四、五任何一处缺失，均报错。",
    ]),
    ("单表关键字段重复检测", [
        "以下四表的对应主键字段禁止重复：表一「工程物料」、表二「物料号」、表四「制造物料」、表五「物料」。",
        "表三「工程物料」允许重复（适配多子件场景）。",
    ]),
    ("项目经理匹配", [
        "表一与表二的「项目经理」必须完全一致。",
        "表一的「工程物料」必须与表二的「物料号」一一对应，其单位、物料类型、项目经理均必须完全一致。",
    ]),
    ("时间逻辑与格式要求", [
        "表一的「生效时间」与表二的「生效时间」必须完全一致。",
        "表一 (生效时间)、表三/表四 (生效日期 & 失效日期) 的填写格式必须为：YYYY-MM-DD HH:MM:SS。",
        "表三及表四中，「失效时间」必须严格晚于「生效时间」。"
    ])
]

y = 1.5
for title, items in rules_2:
    tb = s2.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11.5), Inches(0.4))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = f"● {title}"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    y += 0.45

    for item in items:
        tb_inner = s2.shapes.add_textbox(Inches(0.7), Inches(y), Inches(11.5), Inches(0.35))
        tf_inner = tb_inner.text_frame
        tf_inner.word_wrap = True
        p_i = tf_inner.paragraphs[0]
        p_i.text = f"- {item}"
        p_i.font.size = Pt(14)
        p_i.font.color.rgb = ACCENT
        y += 0.4
    y += 0.15

# Helper to generate a rule card slide
def gen_card_slide(slide_idx, total, title, subtitle, basic_req, forbid_req, cross_rules, extra_logic):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_header_bar(slide, title, subtitle)
    add_slide_number(slide, slide_idx, total)

    col_width = (12.5) / 3
    x_positions = [0.4, 0.4 + col_width, 0.4 + col_width * 2]
    
    def add_box(x, y, w, h, text, bold=False, color=NAVY, size=14, wrap=True):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        return tb

    def add_bullets(x, y, w, h, items, is_bold_header=True):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for item_text, is_header in items:
            p = tf.add_paragraph() if len(tf.paragraphs) > 1 and tf.paragraphs[0].text else tf.paragraphs[0]
            if not p.text: p.text = ""
            if not p.text: 
                p.text = item_text
            else:
                p = tf.add_paragraph()
                p.text = item_text
            
            p.font.size = Pt(12)
            if is_header:
                p.font.bold = True
                p.font.color.rgb = NAVY
                p.space_before = Pt(6)
            else:
                p.font.color.rgb = ACCENT
        return tb

    # Column 1: Required & Forbidden
    items = []
    for t in basic_req: items.append((f"● {t['label']}", True))
    for item in t['items']: items.append((f"  - {item}", False))
    items.append(("", True))
    items.append(("● 禁止填写 (必须留空)", True))
    for t in forbid_req: items.append((f"  - {t['label']}", False))
    
    add_bullets(x_positions[0], 1.5, col_width - 0.1, 5, items)

    # Column 2 & 3: Cross & Extra
    items2 = []
    if cross_rules:
        items2.append(("● 跨表联动规则", True))
        for item in cross_rules: items2.append((f"  - {item}", False))
        items2.append(("", True))
    if extra_logic:
        items2.append(("● 特殊逻辑与校验", True))
        for item in extra_logic: items2.append((f"  - {item}", False))

    add_bullets(x_positions[1], 1.5, col_width * 2 - 0.1, 5, items2)

gen_card_slide(3, TOTAL_SLIDES, "二、表一：BOM 物料清单", 
    subtitle="基础物料定义，作为全系统核心主数据。",
    basic_req=[
        {"label": "必填字段:", "items": ["导入类型, 工程物料, 说明 (禁双空格), 物料类型, 物料组", "工程物料版本, 物料信号, 项目经理", "生效时间 (YYYY-MM-DD HH:MM:SS), 单位, 重量单位"]}
    ],
    forbid_req=[
        {"label": "工程物料版本说明, 材料, 尺寸, 标准, 导入结果", "items": []}
    ],
    cross_rules=[
        "【一致】表一单位/重单 与 表二对应物料必须一致。",
        "【枚举】单位: PCS/MM; 重量单位: KG/G"
    ],
    extra_logic=[
        "【互斥联动】导入类型 = I <-> 物料信号 = GM。",
        "【类型与组映射】类型 1/2/3 对应严格的物料组列表。"
    ]
)

gen_card_slide(4, TOTAL_SLIDES, "三、表二：物料通用数据和客户料号", 
    subtitle="通用数据配置与计划、订货数据设定。",
    basic_req=[
        {"label": "必填字段:", "items": ["导入类型, 物料号, 物料名称 (禁双空格), 物料类型", "物料组, 单位集, 单位, 重量单位, 产品类型", "产品分类, 物料信号, 商品代码, 项目经理", "批次控制 (1 或 2), 保质期/呆滞期, 计划/订货仓库", "计划数据计划员, 订货数据计划员, 订货数据车间计划员", "工艺路线代码、物料代码系统、客户代码、业务伙伴物料代码 (同填或同空)"]}
    ],
    forbid_req=[
        {"label": "原材料, 标准, 大小, 备注 1/2/3, 客户物料开票名称", "items": []},
        {"label": "导入结果, 项目编号, 项目系列", "items": []}
    ],
    cross_rules=[
        "【跨表一致】表二「工程物料」对应的表一「单位、重量单位」必须一致。"
    ],
    extra_logic=[
        "【组校验】工艺代码/系统/客户/伙伴 (4 个)：字段之间互相独立，不再强制同填。",
        "【枚举】物料代码系统若填，必须为 CW。"
    ]
)

gen_card_slide(5, TOTAL_SLIDES, "四、表三：工程版本 BOM 审核", 
    subtitle="BOM 结构定义，重点管控废品率与虚拟件设置。",
    basic_req=[
        {"label": "必填字段 (11 项):", "items": ["导入类型, 工程物料, 版本, 类型, 组件", "净数量, 是否虚拟, 子件仓库, 工序", "生效时间、失效时间 (格式: YYYY-MM-DD HH:MM:SS)"]}
    ],
    forbid_req=[
        {"label": "导入结果", "items": []}
    ],
    cross_rules=[
        "允许「工程物料」在表内重复填写（多子件共用一物料）。"
    ],
    extra_logic=[
        "【废品率 - 优先级1(禁止)】组件含 RR/FL/FLD 或为 PPM000100/0201/0300/0401 -> 空白。",
        "【废品率 - 优先级2(填3.5)】组件含 ABS/COC/PP/PC/PBT... 等。",
        "【废品率 - 优先级3(填2.5)】工程物料或组件倒数两位为 P+数字 或 M+数字。",
        "【废品率 - 优先级4(填1.5)】标准金属牌号，如 SUS301-055 样式。",
        "【虚拟件规则】组件为 PPM000100 或 0300，必须为 1；其余均为 2。",
        "【日期校验】失效时间必须严格晚于生效时间。"
    ]
)

gen_card_slide(6, TOTAL_SLIDES, "五、表四：物料工艺流程", 
    subtitle="工艺路线定义，强校验特征码与工艺描述的一致性。",
    basic_req=[
        {"label": "必填字段 (13 项):", "items": ["导入类型, 制造物料, 工艺流程, 工艺流程说明", "工序, 任务, 工作中心, 机器, 计数点", "生产周期 (分钟)", "生效日期、失效日期 (格式: YYYY-MM-DD HH:MM:SS)", "任务分类"]}
    ],
    forbid_req=[
        {"label": "导入结果", "items": []}
    ],
    cross_rules=[],
    extra_logic=[
        "【固定值字段】「工艺流程」必须固定填写 A10。",
        "【特征码匹配】物料倒数第二位特征码 (S/P/M...) 必须与「工艺流程说明」包含对应词。",
        "【工作中心校验】工作中心 == 机器代码。",
        "【枚举】计数点 只能填写 1 或 2。",
        "【日期校验】失效日期必须严格晚于生效日期。"
    ]
)

gen_card_slide(7, TOTAL_SLIDES, "六、表五：物料扩展属性", 
    subtitle="扩展数据，核心是物料特征码 S/M/P 联动逻辑。",
    basic_req=[
        {"label": "必填字段 (7 项):", "items": ["导入类型, 物料, 产品行业，行业细分", "产品理论重量, 产品应用，项目名称"]}
    ],
    forbid_req=[
        {"label": "物料开票名称, HS 编码, 密度, 料宽, 用金量", "items": []},
        {"label": "机台效率, 模具原因, 麦途物料号, 导入结果", "items": []}
    ],
    cross_rules=[],
    extra_logic=[
        "【跨表存在】表一所有物料必须出现在表五「物料」列。",
        "【联动1】SPM 有值 -> 步距必填。 联动2：电镀方式/类型必须同填。",
        "【行业匹配】「产品行业」与「行业细分」开头的 3 位字符必须一致。",
        "【特征码 S】必填: SPM, PIN, 模号(含 CM), 理论边角料。禁填: 注塑相关。",
        "【特征码 M】必填: 模穴数, 注塑周期, 重量等。禁填: SPM。",
        "【特征码 P】必填: 电镀产速, 电镀方式, 电镀类型。禁填: 模具/注塑。",
    ]
)

# Slide 8: Formulas
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_header_bar(slide, "七、跨表数学公式 (表四 <-> 表五)", "生产周期与其他扩展属性的精确匹配公式。")
add_slide_number(slide, 8, TOTAL_SLIDES)

formulas = [
    ("1. 冲压工艺周期公式", 
     "校验前提：表四工艺流程包含“冲压”且 SPM 已填。\n校验规则：( 生产周期(分钟) / 1.2 ) × SPM 必须等于 1。\n说明：校验生产节拍与标准周期是否匹配。"),
    ("2. 注塑工艺周期公式",
     "校验前提：表四工艺流程包含“成型”或“注塑”。\n校验规则：表四「生产周期」 ≈ ( 注塑周期 / 模穴数 / 60 ) × 1.1。\n说明：根据注塑时间和一模出穴数倒推生产周期，允许极小误差。"),
    ("3. 电镀工艺周期公式",
     "校验前提：表四工艺流程包含“电镀”。\n校验规则：表四「生产周期」 ≈ [ 1 / ( 电镀产速 × 1000 / 步距 ) ] × 1.05。\n说明：基于电镀速度和步距计算生产周期，仅允许极小误差 (5% 或 0.01 分钟)。")
]

y = 1.6
for title, content in formulas:
    shp = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(y), Inches(12.3), Inches(1.4))
    shp.fill.solid()
    shp.fill.fore_color.rgb = LIGHT_BG
    shp.line.color.rgb = NAVY
    
    tb = slide.shapes.add_textbox(Inches(0.7), Inches(y + 0.1), Inches(11.8), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.size = Pt(18)
    p1.font.bold = True
    p1.font.color.rgb = NAVY
    p2 = tf.add_paragraph()
    p2.text = content
    p2.font.size = Pt(13)
    p2.font.color.rgb = ACCENT
    p2.space_before = Pt(4)
    y += 1.65

# Add the new enhancement slide
add_enhancement_slide(9, TOTAL_SLIDES)

prs.save('/workspace/BOM_校验规则详情汇总.pptx')
print("PPT generated successfully.")