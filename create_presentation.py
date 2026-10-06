from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Presentation file name
OUT_FILE = "Ford_After_Sales_2570_2_Slides.pptx"

# Theme colors
FORD_BLUE = RGBColor(0, 52, 128)
FORD_DARK = RGBColor(19, 23, 30)
FORD_WHITE = RGBColor(255, 255, 255)
FORD_LIGHT = RGBColor(230, 236, 244)
FORD_GOLD = RGBColor(208, 164, 56)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


def add_title(slide, text):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(0.45), Inches(8.0), Inches(0.7))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Ford F-1 Light"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = FORD_WHITE
    p.alignment = PP_ALIGN.LEFT
    return box


def add_subtitle(slide, text):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.05), Inches(7.5), Inches(0.5))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Ford F-1 Light"
    p.font.size = Pt(18)
    p.font.color.rgb = FORD_GOLD
    p.alignment = PP_ALIGN.LEFT
    return box


# Slide 1
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
slide1.background.fill.solid()
slide1.background.fill.fore_color.rgb = FORD_BLUE

add_title(slide1, "Ford After-Sales")
add_subtitle(slide1, "Service Improvement 2570")

label = slide1.shapes.add_textbox(Inches(0.6), Inches(1.45), Inches(5.0), Inches(0.4))
label_tf = label.text_frame
p = label_tf.paragraphs[0]
p.text = "Direction: Predictable & Proactive Service"
p.font.name = "Ford F-1 Light"
p.font.size = Pt(12)
p.font.color.rgb = FORD_WHITE

# Left panel
left = slide1.shapes.add_shape(1, Inches(0.6), Inches(1.9), Inches(5.5), Inches(4.0))
left.fill.solid()
left.fill.fore_color.rgb = FORD_LIGHT
left.line.color.rgb = FORD_LIGHT

left_tf = left.text_frame
left_tf.word_wrap = True
left_tf.margin_left = Inches(0.2)
left_tf.margin_top = Inches(0.15)

p = left_tf.paragraphs[0]
p.text = "Key Insight"
p.font.name = "Ford F-1 Light"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = FORD_DARK

for line in [
    "• Visibility",
    "  ลูกค้าและดีลเลอร์ต้องเห็นสถานะแบบเรียลไทม์",
    "",
    "• Predictability",
    "  ต้องให้ ETA / Parts Availability / Capacity แม่นขึ้น",
    "",
    "• Proactive Communication",
    "  ไม่ควรรอให้ลูกค้าโทรมาถาม",
]:
    p = left_tf.add_paragraph()
    p.text = line
    p.font.name = "Ford F-1 Light"
    p.font.size = Pt(14 if "•" in line else 12)
    p.font.color.rgb = FORD_DARK
    if line.startswith("  "):
        p.level = 1

# Right panel
right = slide1.shapes.add_shape(1, Inches(6.35), Inches(1.9), Inches(6.25), Inches(4.0))
right.fill.solid()
right.fill.fore_color.rgb = FORD_WHITE
right.line.color.rgb = FORD_WHITE

right_tf = right.text_frame
right_tf.word_wrap = True
right_tf.margin_left = Inches(0.2)
right_tf.margin_top = Inches(0.15)

p = right_tf.paragraphs[0]
p.text = "Top Initiatives"
p.font.name = "Ford F-1 Light"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = FORD_DARK

items = [
    "1. Automated Service Status Update",
    "   Owner: Head of After-sales Digital",
    "   Timeline: Q1 pilot → Q2 rollout",
    "",
    "2. Parts & ETA Visibility Dashboard",
    "   Owner: Head of Parts",
    "   Timeline: Q1 model → Q3 integrate",
    "",
    "3. Dealer Capacity & Appointment Visibility",
    "   Owner: Head of Dealer Ops",
    "   Timeline: Q1 standard → Q3 expand",
    "",
    "4. Recall Journey Standardization",
    "   Owner: Recall Program Lead",
    "   Timeline: Q1 script & SLA → Q2 VIN dashboard",
]
for line in items:
    p = right_tf.add_paragraph()
    p.text = line
    p.font.name = "Ford F-1 Light"
    p.font.size = Pt(12)
    p.font.color.rgb = FORD_DARK
    if line.startswith("   "):
        p.level = 1

# Quick wins banner
banner = slide1.shapes.add_shape(1, Inches(0.6), Inches(6.15), Inches(12.0), Inches(0.7))
banner.fill.solid()
banner.fill.fore_color.rgb = FORD_GOLD
banner.line.color.rgb = FORD_GOLD
b_tf = banner.text_frame
b_tf.word_wrap = True
p = b_tf.paragraphs[0]
p.text = "Quick Wins: 30 days – SMS/Email Template | 60 days – Standard Script | 90 days – Shared Critical Parts Pool"
p.font.name = "Ford F-1 Light"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = FORD_DARK
p.alignment = PP_ALIGN.CENTER

# Slide 2
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
slide2.background.fill.solid()
slide2.background.fill.fore_color.rgb = FORD_WHITE

add_title = slide2.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(8.0), Inches(0.6))
add_title_tf = add_title.text_frame
p = add_title_tf.paragraphs[0]
p.text = "Action Plan + KPI 2570"
p.font.name = "Ford F-1 Light"
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = FORD_DARK

# Table
rows, cols = 6, 4
shape = slide2.shapes.add_table(rows, cols, Inches(0.5), Inches(1.1), Inches(12.3), Inches(4.8))
table = shape.table
headers = ["Initiative", "Owner", "Timeline", "KPI"]
for c, header in enumerate(headers):
    cell = table.cell(0, c)
    cell.text = header
    for paragraph in cell.text_frame.paragraphs:
        paragraph.font.name = "Ford F-1 Light"
        paragraph.font.size = Pt(11)
        paragraph.font.bold = True
        paragraph.font.color.rgb = FORD_WHITE
    cell.fill.solid()
    cell.fill.fore_color.rgb = FORD_BLUE

records = [
    ["Automated Service Status Update", "Head of After-sales Digital", "Q1 pilot → Q3 rollout", "% cases with auto-status ≥ 90%"],
    ["Parts & ETA Visibility Dashboard", "Head of Parts + Logistics", "Q1 model → Q3 integrate", "Parts fill rate ≥ 95% / ETA accuracy ≥ 85%"],
    ["Dealer Capacity & Appointment Visibility", "Head of Dealer Ops", "Q1 standard → Q3 expand", "On-time completion ≥ 90%"],
    ["Recall Journey Standardization", "SVP Service / Recall Lead", "Q1 script & SLA → Q2 dashboard", "Recall completion within 30 days ≥ 80%"],
    ["Predictive Parts Forecast", "Head of Parts + Data Science", "Q2 POC → Q4 pilot", "Forecast accuracy improves"],
]
for i, row in enumerate(records, start=1):
    for j, value in enumerate(row):
        cell = table.cell(i, j)
        cell.text = value
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.name = "Ford F-1 Light"
            paragraph.font.size = Pt(9)
            paragraph.font.color.rgb = FORD_DARK
        if i % 2 == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(242, 245, 248)

# Decision box
box = slide2.shapes.add_shape(1, Inches(0.5), Inches(6.0), Inches(12.2), Inches(0.9))
box.fill.solid()
box.fill.fore_color.rgb = FORD_BLUE
box.line.color.rgb = FORD_BLUE

btf = box.text_frame
p = btf.paragraphs[0]
p.text = "Decision Required: Approve Q1 pilot budget | Appoint Steering Committee | Approve data access & integration | Launch Quick Wins immediately"
p.font.name = "Ford F-1 Light"
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = FORD_WHITE
p.alignment = PP_ALIGN.CENTER

prs.save(OUT_FILE)
print(f"Created {OUT_FILE}")
