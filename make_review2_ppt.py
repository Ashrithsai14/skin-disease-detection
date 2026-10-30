from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Review_2_Skin_Disease_Detector_v2.pptx"


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*color)


def add_top_bar(slide, title, subtitle=None):
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.72)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(13, 17, 23)
    bar.line.color.rgb = RGBColor(13, 17, 23)

    t = slide.shapes.add_textbox(Inches(0.45), Inches(0.16), Inches(8.5), Inches(0.35))
    p = t.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(125, 211, 252)

    if subtitle:
        s = slide.shapes.add_textbox(Inches(8.8), Inches(0.18), Inches(4.4), Inches(0.3))
        p2 = s.text_frame.paragraphs[0]
        p2.text = subtitle
        p2.alignment = PP_ALIGN.RIGHT
        p2.font.size = Pt(11)
        p2.font.color.rgb = RGBColor(148, 163, 184)


def add_footer(slide, text="Review 2 | Skin Disease Detector"):
    box = slide.shapes.add_textbox(Inches(0.45), Inches(7.05), Inches(12.2), Inches(0.28))
    p = box.text_frame.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.RIGHT
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(100, 116, 139)


def add_title_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, (15, 23, 38))

    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.7), Inches(5.8))
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(22, 29, 43)
    panel.line.color.rgb = RGBColor(36, 50, 74)

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(1.0), Inches(1.3), Inches(0.42))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(37, 99, 235)
    badge.line.color.rgb = RGBColor(37, 99, 235)
    tf = badge.text_frame
    tf.text = "REVIEW 2"
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    title = slide.shapes.add_textbox(Inches(0.95), Inches(1.65), Inches(10.9), Inches(1.3))
    p = title.text_frame.paragraphs[0]
    p.text = "Skin Disease Detector"
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = RGBColor(229, 238, 252)

    sub = slide.shapes.add_textbox(Inches(0.95), Inches(2.55), Inches(10.5), Inches(0.8))
    tf = sub.text_frame
    p = tf.paragraphs[0]
    p.text = "A deep learning web application that classifies skin images and shows supportive health guidance."
    p.font.size = Pt(18)
    p.font.color.rgb = RGBColor(148, 163, 184)

    names = slide.shapes.add_textbox(Inches(0.95), Inches(3.15), Inches(10.7), Inches(0.8))
    tf = names.text_frame
    p = tf.paragraphs[0]
    p.text = "Ashrith Sai (23BCE8770)    |    Ruturaj (23BCE8871)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(226, 232, 240)

    left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(3.35), Inches(5.45), Inches(1.25))
    left.fill.solid()
    left.fill.fore_color.rgb = RGBColor(15, 34, 59)
    left.line.color.rgb = RGBColor(47, 79, 122)
    tx = left.text_frame
    tx.word_wrap = True
    tx.text = "Project Type\nDeep Learning + Flask Web App"
    tx.paragraphs[0].font.size = Pt(14)
    tx.paragraphs[0].font.bold = True
    tx.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    tx.paragraphs[1].font.size = Pt(12)
    tx.paragraphs[1].font.color.rgb = RGBColor(203, 213, 225)

    right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.55), Inches(3.35), Inches(5.45), Inches(1.25))
    right.fill.solid()
    right.fill.fore_color.rgb = RGBColor(15, 34, 59)
    right.line.color.rgb = RGBColor(47, 79, 122)
    tx = right.text_frame
    tx.word_wrap = True
    tx.text = "Main Output\nPrediction, disease details, advice, PDF report"
    tx.paragraphs[0].font.size = Pt(14)
    tx.paragraphs[0].font.bold = True
    tx.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    tx.paragraphs[1].font.size = Pt(12)
    tx.paragraphs[1].font.color.rgb = RGBColor(203, 213, 225)

    add_footer(slide)


def add_bullets_slide(title, subtitle, bullets, box_title="Key Points"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, (15, 23, 38))
    add_top_bar(slide, "Skin Disease Detector", subtitle)
    add_footer(slide)

    ttl = slide.shapes.add_textbox(Inches(0.6), Inches(0.95), Inches(12), Inches(0.55))
    p = ttl.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(229, 238, 252)

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(1.6), Inches(11.8), Inches(5.1))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(22, 29, 43)
    card.line.color.rgb = RGBColor(36, 50, 74)

    head = slide.shapes.add_textbox(Inches(1.0), Inches(1.82), Inches(3.5), Inches(0.3))
    p = head.text_frame.paragraphs[0]
    p.text = box_title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(125, 211, 252)

    tx = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.0), Inches(4.0))
    tf = tx.text_frame
    tf.word_wrap = True
    tf.clear()

    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = bullet
        p.level = 0
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(226, 232, 240)
        p.space_after = Pt(10)


def add_workflow_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, (15, 23, 38))
    add_top_bar(slide, "Skin Disease Detector", "Workflow")
    add_footer(slide)

    ttl = slide.shapes.add_textbox(Inches(0.6), Inches(0.95), Inches(12), Inches(0.55))
    p = ttl.text_frame.paragraphs[0]
    p.text = "System Workflow"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(229, 238, 252)

    steps = [
        "1. User uploads a skin image",
        "2. Frontend sends it to Flask backend",
        "3. Image is resized and normalized",
        "4. CNN model predicts the skin class",
        "5. Result and advice are displayed",
        "6. PDF report can be generated and shared",
    ]

    left = 0.85
    top = 1.8
    w = 3.55
    h = 1.0
    gap = 0.25

    for idx, step in enumerate(steps):
        row = idx // 2
        col = idx % 2
        x = Inches(left + col * (w + gap))
        y = Inches(top + row * (h + 0.35))
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(w), Inches(h))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(15, 34, 59)
        box.line.color.rgb = RGBColor(47, 79, 122)
        tf = box.text_frame
        tf.word_wrap = True
        tf.text = step
        p = tf.paragraphs[0]
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    note = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(5.35), Inches(11.5), Inches(1.1))
    note.fill.solid()
    note.fill.fore_color.rgb = RGBColor(20, 78, 113)
    note.line.color.rgb = RGBColor(56, 189, 248)
    tf = note.text_frame
    tf.text = "The model gives a preliminary prediction only. Final diagnosis must be confirmed by a dermatologist."
    p = tf.paragraphs[0]
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)


def add_table_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, (15, 23, 38))
    add_top_bar(slide, "Skin Disease Detector", "Model and Stack")
    add_footer(slide)

    ttl = slide.shapes.add_textbox(Inches(0.6), Inches(0.95), Inches(12), Inches(0.55))
    p = ttl.text_frame.paragraphs[0]
    p.text = "Model, Technology, and Features"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(229, 238, 252)

    left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(1.6), Inches(5.8), Inches(4.9))
    left.fill.solid()
    left.fill.fore_color.rgb = RGBColor(22, 29, 43)
    left.line.color.rgb = RGBColor(36, 50, 74)
    tx = left.text_frame
    tx.word_wrap = True
    tx.text = "Model Details\n\n- Trained CNN-based classifier\n- Uses ResNet preprocessing\n- Predicts one of 10 skin classes\n- Saved model file used: model.h5"
    for i, para in enumerate(tx.paragraphs):
        para.font.size = Pt(15 if i == 0 else 13)
        para.font.bold = i == 0
        para.font.color.rgb = RGBColor(255, 255, 255) if i == 0 else RGBColor(226, 232, 240)

    right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.75), Inches(1.6), Inches(5.8), Inches(4.9))
    right.fill.solid()
    right.fill.fore_color.rgb = RGBColor(22, 29, 43)
    right.line.color.rgb = RGBColor(36, 50, 74)
    tx = right.text_frame
    tx.word_wrap = True
    tx.text = "Technology Stack\n\n- Python and Flask\n- TensorFlow / Keras\n- HTML, CSS, JavaScript\n- PIL, NumPy, ReportLab\n- Firebase and Twilio integrations"
    for i, para in enumerate(tx.paragraphs):
        para.font.size = Pt(15 if i == 0 else 13)
        para.font.bold = i == 0
        para.font.color.rgb = RGBColor(255, 255, 255) if i == 0 else RGBColor(226, 232, 240)


def add_conclusion_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, (15, 23, 38))
    add_top_bar(slide, "Skin Disease Detector", "Conclusion")
    add_footer(slide)

    ttl = slide.shapes.add_textbox(Inches(0.6), Inches(0.95), Inches(12), Inches(0.55))
    p = ttl.text_frame.paragraphs[0]
    p.text = "Conclusion and Future Scope"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(229, 238, 252)

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(1.6), Inches(11.8), Inches(4.9))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(22, 29, 43)
    card.line.color.rgb = RGBColor(36, 50, 74)

    tx = slide.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.0), Inches(4.2))
    tf = tx.text_frame
    tf.word_wrap = True
    bullets = [
        "This project shows how deep learning can support early skin condition screening.",
        "It is a helpful assistant tool, not a replacement for a medical diagnosis.",
        "The app combines prediction, reporting, location help, and patient guidance in one interface.",
        "Future scope: train on a larger medical dataset, improve accuracy, and add more robust deployment support.",
    ]
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(226, 232, 240)
        p.space_after = Pt(12)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

add_title_slide()
add_bullets_slide(
    "Problem Statement",
    "Why this project matters",
    [
        "Skin diseases are common and can be difficult to identify quickly without expert help.",
        "A fast preliminary screening tool can help users understand what condition may be present.",
        "The project provides a simple image-based interface for this first-level analysis.",
    ],
)
add_bullets_slide(
    "Objective",
    "What the system is designed to do",
    [
        "Classify uploaded skin images using a CNN-based deep learning model.",
        "Show the predicted condition with treatment and diet information.",
        "Generate a PDF report and support doctor search and WhatsApp sharing.",
    ],
    box_title="Project Goals",
)
add_workflow_slide()
add_table_slide()
add_bullets_slide(
    "Supported Classes",
    "Model output categories",
    [
        "Eczema",
        "Melanoma",
        "Atopic Dermatitis",
        "Basal Cell Carcinoma",
        "Melanocytic Nevi",
        "Benign Keratosis-like Lesions",
        "Psoriasis",
        "Seborrheic Keratoses",
        "Tinea / Fungal Infections",
        "Warts / Viral Infections",
    ],
    box_title="Classes",
)
add_conclusion_slide()

prs.save(OUT)
print(f"Saved: {OUT}")
