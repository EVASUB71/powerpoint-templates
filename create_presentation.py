from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUTPUT = Path(__file__).with_name("Business_Review_Presentation.pptx")

NAVY = "003366"
GREEN = "4CAF50"
ORANGE = "FF9800"
BLUE = "2196F3"
YELLOW = "FFD54F"
INK = "18324B"
MUTED = "66788A"
PALE = "F3F6F9"
WHITE = "FFFFFF"
FONT = "Aptos"


def rgb(value):
    return RGBColor.from_string(value)


def add_text(slide, text, x, y, w, h, size=14, color=INK, bold=False,
             align=PP_ALIGN.LEFT, font=FONT, margin=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def add_rect(slide, x, y, w, h, fill, radius=False, line=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.fill.background() if line is None else None
    if line:
        shape.line.color.rgb = rgb(line)
    if radius:
        shape.adjustments[0] = 0.12
    return shape


def add_header(slide, title, subtitle):
    add_rect(slide, 0, 0, 13.333, 1.23, NAVY)
    add_text(slide, title, 0.62, 0.19, 12.1, 0.48, 25, WHITE, True)
    add_text(slide, subtitle, 0.65, 0.72, 12.0, 0.28, 11, "D9E6F2", True)


def add_section_heading(slide, text, x, y, w):
    add_rect(slide, x, y, w, 0.38, NAVY, radius=True)
    add_text(slide, text, x + 0.16, y + 0.01, w - 0.32, 0.34, 11, WHITE, True)


def add_badge(slide, glyph, x, y, color=NAVY, diameter=0.38):
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(diameter), Inches(diameter)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = rgb(PALE)
    circle.line.fill.background()
    add_text(slide, glyph, x, y - 0.005, diameter, diameter, 16, color, True,
             PP_ALIGN.CENTER, font="Segoe UI Symbol")


def add_footer(slide, message):
    add_rect(slide, 0.55, 6.82, 12.23, 0.54, PALE, radius=True)
    add_rect(slide, 0.55, 6.82, 1.48, 0.54, NAVY, radius=True)
    add_text(slide, "KEY MESSAGE", 0.69, 6.84, 1.2, 0.48, 9, WHITE, True)
    add_text(slide, message, 2.23, 6.84, 10.25, 0.48, 11, INK)


def build_slide_one(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "MID-YEAR BUSINESS REVIEW",
               "BUSINESS DEVELOPMENT & SALES H1 2026")

    add_section_heading(slide, "SALES PERFORMANCE", 0.55, 1.48, 6.0)
    add_rect(slide, 0.55, 1.94, 6.0, 3.35, WHITE, radius=True, line="DCE4EB")

    metrics = [
        ("◎", "Annual Target", "€200k", NAVY),
        ("▥", "Current Sales", "€81k", GREEN),
        ("↗", "Expected by June Close", "€100k", NAVY),
        ("◷", "Achievement", "50% of FY Target", ORANGE),
    ]
    for i, (glyph, label, value, value_color) in enumerate(metrics):
        y = 2.12 + i * 0.76
        add_badge(slide, glyph, 0.82, y + 0.08)
        add_text(slide, label, 1.38, y + 0.03, 3.1, 0.27, 12, MUTED)
        add_text(slide, value, 1.38, y + 0.32, 3.1, 0.31, 16, value_color, True)
        if i < 3:
            add_rect(slide, 1.38, y + 0.71, 4.82, 0.012, "E8EDF2")

    chart_data = CategoryChartData()
    chart_data.categories = ["Achieved", "Remaining"]
    chart_data.add_series("FY Target", (50, 50))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(4.88), Inches(2.76), Inches(1.52), Inches(1.52),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.has_title = False
    chart.chart_style = 10
    chart.plots[0].hole_size = 74
    chart.plots[0].vary_by_categories = True
    chart.series[0].points[0].format.fill.solid()
    chart.series[0].points[0].format.fill.fore_color.rgb = rgb(NAVY)
    chart.series[0].points[1].format.fill.solid()
    chart.series[0].points[1].format.fill.fore_color.rgb = rgb("DCE4EB")
    chart.series[0].points[0].format.line.fill.background()
    chart.series[0].points[1].format.line.fill.background()
    add_text(slide, "50%", 5.07, 3.28, 1.14, 0.42, 20, NAVY, True, PP_ALIGN.CENTER)
    add_text(slide, "OF FY TARGET", 4.97, 3.72, 1.34, 0.23, 7, MUTED, True,
             PP_ALIGN.CENTER)

    add_section_heading(slide, "MY FOCUS DURING H1", 6.83, 1.48, 5.95)
    add_rect(slide, 6.83, 1.94, 5.95, 3.35, WHITE, radius=True, line="DCE4EB")
    focus = [
        ("●", "Deliver personal sales"),
        ("▽", "Build a qualified pipeline"),
        ("↔", "Develop strategic partnerships"),
        ("✦", "Generate new lead channels"),
        ("↗", "Support sustainable country growth"),
    ]
    for i, (glyph, text) in enumerate(focus):
        y = 2.15 + i * 0.59
        add_badge(slide, glyph, 7.12, y + 0.03, diameter=0.36)
        add_text(slide, text, 7.68, y, 4.72, 0.42, 12, INK, i == 0)

    add_footer(
        slide,
        "Balancing short-term revenue delivery with long-term business development.",
    )
    return slide


def build_slide_two(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "H2 SALES PIPELINE", "QUALIFIED SALES PIPELINE")

    add_section_heading(slide, "ACTIVE OPPORTUNITIES", 0.55, 1.48, 7.65)
    table_x, table_y = 0.55, 1.96
    widths = (4.55, 1.35, 1.75)
    headers = ("STAGE", "NOK", "EUR (APPROX.)")
    x = table_x
    for header, width in zip(headers, widths):
        add_rect(slide, x, table_y, width, 0.43, NAVY)
        add_text(slide, header, x + 0.16, table_y + 0.02, width - 0.3, 0.38,
                 9, WHITE, True,
                 PP_ALIGN.LEFT if header == "STAGE" else PP_ALIGN.CENTER)
        x += width

    rows = [
        ("01 - Prospecting", "185k", "€16k", "FFF4CC"),
        ("02 - Interest Expressed", "490k", "€43k", YELLOW),
        ("03 - Bid in Progress", "217k", "€19k", ORANGE),
        ("05 - Quote Sent", "175k", "€15k", BLUE),
    ]
    for row_index, (stage, nok, eur, row_color) in enumerate(rows):
        y = table_y + 0.43 + row_index * 0.63
        fill = WHITE if row_index % 2 == 0 else PALE
        x = table_x
        for col_index, (value, width) in enumerate(zip((stage, nok, eur), widths)):
            add_rect(slide, x, y, width, 0.62, fill, line="E5EBF0")
            if col_index == 0:
                add_rect(slide, x + 0.13, y + 0.12, 0.1, 0.38, row_color)
            add_text(slide, value, x + (0.34 if col_index == 0 else 0.08), y + 0.04,
                     width - (0.42 if col_index == 0 else 0.16), 0.52,
                     12, INK, col_index > 0,
                     PP_ALIGN.LEFT if col_index == 0 else PP_ALIGN.CENTER)
            x += width

    add_rect(slide, 8.58, 1.48, 4.2, 2.25, NAVY, radius=True)
    add_text(slide, "TOTAL ACTIVE PIPELINE", 8.9, 1.78, 3.55, 0.33,
             11, "D9E6F2", True, PP_ALIGN.CENTER)
    add_text(slide, "≈ €93k", 8.83, 2.21, 3.7, 0.72, 31, WHITE, True,
             PP_ALIGN.CENTER)
    add_text(slide, "(=1.07 MNOK)", 8.9, 3.01, 3.55, 0.32,
             12, "D9E6F2", True, PP_ALIGN.CENTER)

    add_section_heading(slide, "PIPELINE HIGHLIGHTS", 0.55, 5.02, 12.23)
    highlights = [
        "Strong opportunities entering H2",
        "Healthy mix of early and advanced opportunities",
        "Several projects already in quotation/bid phase",
    ]
    for i, text in enumerate(highlights):
        x = 0.78 + i * 4.08
        add_badge(slide, "✓", x, 5.65, GREEN, 0.4)
        add_text(slide, text, x + 0.55, 5.59, 3.35, 0.55, 11, INK)

    add_footer(
        slide,
        "A balanced pipeline supporting revenue generation throughout H2.",
    )
    return slide


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Mid-Year Business Review and H2 Sales Pipeline"
    prs.core_properties.subject = "Editable H1 2026 business review presentation"
    build_slide_one(prs)
    build_slide_two(prs)
    prs.save(OUTPUT)
    print(f"Created {OUTPUT}")


if __name__ == "__main__":
    main()
