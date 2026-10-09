"""Generate a business review using editable PowerPoint objects."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


# Monthly content: edit this section, not the rendering functions below.
OUTPUT_FILE = Path(__file__).resolve().parent / "Business_Review_Presentation.pptx"
PERIOD = "H1 2026"
COLORS = {
    "navy": "003366",
    "green": "4CAF50",
    "orange": "FF9800",
    "blue": "2196F3",
    "yellow": "FFD54F",
    "purple": "7E57C2",
    "dark_purple": "5E408D",
    "light_green": "80B650",
    "ink": "263B50",
    "muted": "60758A",
    "background": "F4F7FA",
    "white": "FFFFFF",
    "border": "DFE7EE",
}
EVENTS = [
    "Mobile World Congress (MWC)",
    "Nit de les Telecomunicacions",
    "ENERTIC IA",
    "I2CAT Aerospace & Defence",
    "Meeting d'Auto de CIAC",
    "IT Revista TIC",
]
CONSULTANTS = [
    ("NEW", "AUIKEN", "Prospecting specialized consultants and market influencers"),
    ("NEW", "BABEL", "Prospecting specialized consultants and market influencers"),
    ("NEW", "ASEITEC", "Prospecting specialized consultants and market influencers"),
    ("EXPANSION", "DERTEN TISAX", "Building scalable lead generation channels"),
]
COUNTRIES = ["Germany", "Finland", "France", "Sweden", "Denmark"]
SECTORS = [
    "Nuclear", "ICT & AI", "Automotive", "Health",
    "Defence & Aerospace", "Innovation Management",
]
SLIDES = [
    {
        "title": "Mid-Year Business Review",
        "subtitle": "Sales performance and foundations for sustainable growth",
        "metrics": [
            ("€200k", "Annual Target"),
            ("€81k", "Current Sales"),
            ("€100k", "Expected by June Close"),
            ("50%", "Achievement"),
        ],
        "columns": 2,
        "cards": [
            {
                "title": "SALES PERFORMANCE", "color": "blue",
                "lines": [
                    "Deliver personal sales and convert near-term opportunities.",
                    "Maintain a qualified pipeline to support the second half.",
                    "Keep country contribution and future growth in balance.",
                ],
            },
            {
                "title": "FIVE FOCUS ITEMS", "color": "green",
                "lines": [
                    "① Deliver personal sales",
                    "② Build qualified pipeline",
                    "③ Develop strategic partnerships",
                    "④ Generate new lead channels",
                    "⑤ Support sustainable country growth",
                ],
            },
        ],
        "footer": "BALANCING SHORT-TERM SALES DELIVERY WITH LONG-TERM, SUSTAINABLE COUNTRY GROWTH",
    },
    {
        "title": "H2 Sales Pipeline",
        "subtitle": "Turning qualified opportunities into second-half sales",
        "metrics": [("€93k", "Total Active Pipeline"), ("1.07 MNOK", "Total Active Pipeline")],
        "columns": 2,
        "cards": [
            {
                "title": "ACTIVE PIPELINE BY STAGE", "color": "blue",
                "table": {
                    "headers": ["Stage", "NOK", "EUR"],
                    "rows": [
                        ["Prospecting", "185k", "€16k"],
                        ["Interest Expressed", "490k", "€43k"],
                        ["Bid in Progress", "217k", "€19k"],
                        ["Quote Sent", "175k", "€15k"],
                    ],
                    "widths": [0.54, 0.23, 0.23],
                },
            },
            {
                "title": "PIPELINE HIGHLIGHTS", "color": "green",
                "lines": [
                    "✓ Active opportunities across all four sales stages",
                    "✓ Progress qualified interest into bids and quotations",
                    "✓ Prioritize conversion while continuing pipeline development",
                ],
            },
        ],
        "footer": "BUILDING PIPELINE QUALITY AND CONVERSION TO SUPPORT H2 SALES DELIVERY",
    },
    {
        "title": "Niches & Industries",
        "subtitle": "Focusing on strategic sectors to drive future growth",
        "metrics": [
            ("5", "Strategic Sectors"), ("19", "Leads Generated"),
            ("€25.5k", "Revenue (Training)"), ("2", "Services Under Development"),
        ],
        "columns": 6,
        "cards": [
            {
                "title": "NUCLEAR", "color": "green",
                "lines": [
                    "ISO 19443",
                    "Objective: 3 opportunities",
                    "Focus: ongoing ecosystem development and market positioning",
                    "Current status: 1 lead generated",
                ],
            },
            {
                "title": "ICT & AI", "color": "blue",
                "lines": [
                    "ISO 42001",
                    "Private training for BABEL: €6k (September)",
                    "Webinar with ASEITEC",
                    "Prospecting in progress: target 2 opportunities",
                    "Next step: September webinar and continued prospecting",
                ],
            },
            {
                "title": "AUTOMOTIVE", "color": "purple",
                "lines": [
                    "TISAX",
                    "Growing with specialist consultant DERTEN",
                    "Partnered with CIAC ecosystem",
                    "Attended Meeting d'Auto",
                    "11 leads generated",
                    "Recovery of KC Zanini Global: €5.5k",
                ],
            },
            {
                "title": "HEALTH", "color": "dark_purple",
                "lines": [
                    "ISO 7101",
                    "1 opportunity identified",
                    "Coordination with NSD Committee",
                    "Collaboration with Italy and Brazil",
                    "Sector support with Silvia, Ainhoa & Natalia",
                    "New service under development",
                ],
            },
            {
                "title": "DEFENCE & AEROSPACE", "color": "navy",
                "lines": [
                    "PART IS",
                    "New service under development with NSD Committee",
                    "Build capabilities and position DNV as a trusted partner",
                    "Strengthen engagement and create sector opportunities",
                ],
            },
            {
                "title": "INNOVATION MANAGEMENT", "color": "green",
                "lines": [
                    "ISO 56001",
                    "High growth potential standard",
                    "First project sold: Moeve Aviation, €2k",
                    "Supporting innovation and business transformation",
                    "Key customer project: Moeve Aviation",
                    "ISO 56001 certified",
                ],
            },
        ],
        "footer": "BUILDING EXPERTISE, VISIBILITY AND COMMERCIAL TRACTION FOR MEDIUM- AND LONG-TERM GROWTH IN SPAIN",
        "footer_detail": "✓ Creating market relevance    ✓ Generating leads and opportunities    ✓ Building foundations for future growth",
    },
    {
        "title": "Network & External Communication",
        "subtitle": "Building market visibility and expanding influence ecosystems",
        "metrics": [
            ("6 / 3", "Industry Events · 200%"),
            ("4 / 7", "Active Consultants · 57%"),
            ("0 / 3", "Certification Ceremonies · 0%"),
        ],
        "columns": 3,
        "cards": [
            {
                "title": "INDUSTRY VISIBILITY", "color": "blue",
                "lines": [
                    "Events & market presence",
                    *["✓ " + event for event in EVENTS],
                    "Impact: visibility in AI, ICT, Automotive, Aerospace, Defence and Digital Transformation",
                    "New relationships supporting future lead generation",
                ],
            },
            {
                "title": "CERTIFICATION CEREMONIES", "color": "orange",
                "lines": [
                    "Objective: 3 certification ceremonies in 2026",
                    "Opportunity identified with TUBACEX",
                    "Final presentation assigned internally to another colleague",
                    "Customer relationship maintained",
                    "Current status: 0 ceremonies delivered in H1",
                ],
            },
            {
                "title": "CONSULTANT ECOSYSTEM", "color": "green",
                "lines": ["Objective: 7 active consultants generating leads"],
                "table": {
                    "headers": ["Status", "Partner", "Ongoing Activity"],
                    "rows": CONSULTANTS,
                    "widths": [0.24, 0.23, 0.53],
                },
            },
        ],
        "footer": "BUILDING VISIBILITY, INFLUENCE AND REFERRAL ECOSYSTEMS FOR GROWTH IN STRATEGIC SECTORS",
        "footer_detail": "✓ Increase market presence    ✓ Strengthen relationships and influence    ✓ Expand the consultant network",
    },
    {
        "title": "Business Development",
        "subtitle": "Building a scalable growth engine through partnerships",
        "metrics": [
            ("€45k", "Closed Revenue · 2024"),
            ("16", "Partners · 3 established / 1 high-potential / 12 developing"),
            ("~€100k", "Qualified Leads · BABEL"),
            ("€16k", "Active Pipeline · AUIKEN + ASEITEC"),
        ],
        "columns": 4,
        "cards": [
            {
                "title": "BABEL", "color": "blue",
                "lines": [
                    "KEY GROWTH OPPORTUNITY",
                    "High-potential partner",
                    "Our biggest opportunity for growth in AI & Digital Transformation",
                    "€5k closed revenue delivered in 2024",
                    "~€100k qualified lead potential",
                    "Focus: innovation and digital solutions",
                    "H2 priority: cross-functional collaboration in progress",
                ],
            },
            {
                "title": "AUIKEN", "color": "green",
                "lines": [
                    "Established partner",
                    "€12k closed revenue",
                    "€10k active pipeline",
                    "Key opportunity: Banco Santander",
                ],
            },
            {
                "title": "ASEITEC", "color": "orange",
                "lines": [
                    "Established partner",
                    "€10k closed revenue",
                    "€6k active pipeline",
                    "Current focus: ISO 42001 September webinar",
                ],
            },
            {
                "title": "DERTEN", "color": "purple",
                "lines": [
                    "Established partner",
                    "€9k revenue in 2025",
                    "€18k revenue in 2026",
                    "+100% year-on-year growth",
                    "Specialized TISAX consultant",
                ],
            },
        ],
        "footer": "DEVELOPING PARTNERSHIPS THAT DELIVER REVENUE TODAY AND SCALABLE OPPORTUNITIES TOMORROW",
    },
    {
        "title": "Regional & Cross-Unit Collaboration",
        "subtitle": "Leveraging DNV global capabilities to accelerate growth in Iberia",
        "metrics": [
            ("5", "Countries · " + ", ".join(COUNTRIES)),
            ("2", "Business Units · DNV Energy, DNV Maritime"),
            ("4", "Opportunity Cards Developed"),
            ("1", "Iberia Market Representative"),
        ],
        "columns": 4,
        "cards": [
            {
                "title": "INTERNATIONAL COLLABORATION", "color": "blue",
                "lines": [
                    "Countries: " + ", ".join(COUNTRIES),
                    "✓ Access to technical expertise and best practices",
                    "✓ Market intelligence exchange",
                    "✓ Support for customer opportunities in Spain",
                    "✓ Iberia representation in international business development discussions",
                    "Market trends, emerging sectors and growth opportunities",
                ],
            },
            {
                "title": "CROSS BUSINESS COLLABORATION", "color": "green",
                "lines": [
                    "DNV ENERGY",
                    "Moeve tender and Moeve Aviation opportunity",
                    "Revenue contribution: €4k",
                    "DNV MARITIME",
                    "Connection with Norwegian Chamber of Commerce",
                    "Invitation to Bilbao Maritime and Fisheries event",
                    "Joint cross-selling initiatives planned for September 2026",
                ],
            },
            {
                "title": "NEW SERVICE ENABLEMENT", "color": "purple",
                "lines": [
                    "✓ ISO 7101 (Health): approved Iberia business case",
                    "✓ PART IS (Defence & Aerospace): approved Iberia business case",
                    "✓ ISO 56001: approved; first project sold (€2k)",
                    "✗ UNE 158: market analysis supported a no-go decision",
                    "NSD Committee: card preparation and presentation, market sizing, financial viability and alignment with technical experts",
                ],
            },
            {
                "title": "BUSINESS IMPACT", "color": "orange",
                "lines": [
                    "GROWTH ENABLED",
                    "✓ ISO 7101 (Health)",
                    "NEW CAPABILITY ENABLED",
                    "✓ PART IS (Defence & Aerospace)",
                    "FIRST COMMERCIAL VALIDATION",
                    "✓ ISO 56001: €2k project sold",
                    "INVESTMENT AVOIDED",
                    "✓ UNE 158: evidence-based no-go decision",
                ],
            },
        ],
        "footer": "CONNECTING INTERNATIONAL EXPERTISE, BUSINESS UNITS AND BUSINESS CASES FOR BETTER GROWTH AND PORTFOLIO DECISIONS",
        "footer_detail": "Supporting 3 new service introductions, 2 cross-business initiatives and collaboration across 5 countries",
    },
    {
        "title": "Roadmap Execution",
        "subtitle": "Contributing to Iberia strategic priorities through business development",
        "metrics": [
            ("€92k", "Revenue Contribution · H1"),
            ("3", "New Services Enabled"),
            ("Active", "Customer Support Model"),
            ("6", "Strategic Sectors Developed"),
            ("€22k+", "Cross Selling Revenue"),
        ],
        "columns": 5,
        "cards": [
            {
                "title": "GROWTH", "color": "green",
                "lines": [
                    "KEY DRIVER",
                    "Delivering today's revenue while building foundations for sustainable growth",
                    "€92k revenue delivered in H1",
                    "€92k active pipeline entering H2",
                    "Balance short-term sales with long-term growth initiatives",
                ],
            },
            {
                "title": "CAPACITY BUILDING FOR NEW SCHEMES", "color": "purple",
                "lines": [
                    "KEY ENABLER",
                    "Expand capability to deliver high-value solutions in growing markets",
                    "4 opportunity cards developed",
                    "3 new services: ISO 7101, PART IS, ISO 56001",
                    "1 investment avoided: UNE 158",
                    "Coordination with NSD Committee",
                ],
            },
            {
                "title": "SUPERIOR CUSTOMER EXPERIENCE", "color": "blue",
                "lines": [
                    "KEY FOCUS",
                    "Support customers throughout their journey for long-term value and satisfaction",
                    "Customer meetings when requested",
                    "Follow operational implementation and evolution",
                    "Support complex and cross-business opportunities",
                    "Access to international expertise and new capabilities",
                ],
            },
            {
                "title": "INDUSTRY FOCUS", "color": "orange",
                "lines": [
                    "KEY IMPACT",
                    "Position DNV in high-potential sectors to capture future growth",
                    *["● " + sector for sector in SECTORS],
                ],
            },
            {
                "title": "CROSS SELLING", "color": "light_green",
                "lines": [
                    "KEY VALUE",
                    "Leverage internal and auditor collaboration across geographies",
                    "€22k+ cross-selling contribution in H1",
                    "€11k Bestinver",
                    "Internal and auditor collaboration model",
                ],
            },
        ],
        "footer": "DELIVERING TODAY'S REVENUE WHILE ENABLING TOMORROW'S GROWTH CAPABILITIES",
    },
    {
        "title": "Business Development Actions & Lead Generation",
        "subtitle": "Building scalable lead generation channels for Iberia growth",
        "metrics": [
            ("35 / 80", "Leads Shared with Sales"),
            ("4 / 7", "Active Consultants"),
            ("6 / 3", "Industry Events"),
            ("35+", "Qualified Leads Generated"),
            ("3", "Auditor Collaboration Initiatives"),
            ("80", "Annual Target to Sales Team"),
        ],
        "columns": 4,
        "cards": [
            {
                "title": "CONSULTANT CHANNEL", "color": "green",
                "lines": [
                    "AUIKEN · ASEITEC · BABEL · DERTEN",
                    "€45k revenue generated",
                    "€16k active pipeline",
                    "~€100k qualified lead potential",
                    "Strong channel to access new accounts and opportunities",
                ],
            },
            {
                "title": "INDUSTRY VISIBILITY", "color": "purple",
                "lines": [
                    *["✓ " + event for event in EVENTS],
                    "Visibility and relationships through key industry events and media",
                ],
            },
            {
                "title": "INDUSTRY ECOSYSTEMS", "color": "blue",
                "lines": [
                    "✓ Active membership in CIAC",
                    "✓ Access to automotive ecosystem and key decision makers",
                    "✓ Support for TISAX growth strategy in Iberia",
                    "Open doors and generate quality opportunities",
                ],
            },
            {
                "title": "AUDITOR COLLABORATION", "color": "orange",
                "lines": [
                    "SPAIN: 2 projects generated; ~€7k revenue",
                    "PORTUGAL: 2 projects won; additional opportunities in pipeline",
                    "Cross-selling through auditor relationships is becoming a scalable lead generation model for Iberia",
                ],
            },
        ],
        "comparison": [
            {
                "title": "H1 REALITY", "color": "orange",
                "lines": [
                    "Country contribution requirements increased focus on lead conversion during Q2 to support country revenue.",
                    "✓ Business development initiatives progressed in parallel across all channels.",
                ],
            },
            {
                "title": "H2 FOCUS", "color": "blue",
                "lines": [
                    "✓ Scale consultant channel and expand partnerships.",
                    "✓ Strengthen auditor collaboration in Spain and Portugal.",
                    "✓ Increase lead transfer to Sales teams (target: 80 leads).",
                    "✓ Maximize presence in industry events and ecosystems.",
                ],
            },
        ],
        "footer": "BALANCING SHORT-TERM COUNTRY CONTRIBUTION WITH THE DEVELOPMENT OF SCALABLE LEAD GENERATION MECHANISMS FOR IBERIA",
    },
]


# Layout only: all business values and copy are configured above.
SLIDE_WIDTH = 16
SLIDE_HEIGHT = 9
MARGIN = 0.45
GAP = 0.18
FONT = "Aptos"


def color(name):
    return RGBColor.from_string(COLORS.get(name, name))


def box(slide, x, y, width, height, fill, border=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color(fill)
    if border:
        shape.line.color.rgb = color(border)
    else:
        shape.line.fill.background()
    return shape


def write(frame, lines, size=14, ink="ink", bold=False, align=PP_ALIGN.LEFT,
          spacing=6):
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Inches(0.05)
    frame.margin_top = frame.margin_bottom = Inches(0.03)
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.alignment = align
        paragraph.font.name = FONT
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color(ink)
        paragraph.space_after = Pt(spacing)
        paragraph.line_spacing = 1.08


def text(slide, x, y, width, height, lines, **style):
    shape = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(width), Inches(height)
    )
    write(shape.text_frame, lines, **style)
    return shape


def table(slide, data, x, y, width, height):
    rows = data["rows"]
    shape = slide.shapes.add_table(
        len(rows) + 1, len(data["headers"]),
        Inches(x), Inches(y), Inches(width), Inches(height),
    )
    grid = shape.table
    for column, fraction in zip(grid.columns, data["widths"]):
        column.width = Inches(width * fraction)
    for row_index, values in enumerate([data["headers"], *rows]):
        for column_index, value in enumerate(values):
            cell = grid.cell(row_index, column_index)
            cell.fill.solid()
            cell.fill.fore_color.rgb = color(
                "navy" if row_index == 0 else
                ("background" if row_index % 2 else "white")
            )
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            write(
                cell.text_frame, [value], size=10 if width < 6 else 12,
                ink="white" if row_index == 0 else "ink",
                bold=row_index == 0,
            )
    return shape


def card(slide, data, x, y, width, height, compact=False):
    box(slide, x, y, width, height, "white", "border")
    header_height = 0.48 if height < 2 else (0.88 if width < 3.2 else 0.68)
    box(slide, x, y, width, header_height, data["color"])
    text(
        slide, x + 0.12, y + 0.08, width - 0.24, header_height - 0.12,
        [data["title"]], size=13 if width < 3.2 else 15,
        ink="white", bold=True,
    )
    body_y = y + header_height + 0.14
    body_height = height - header_height - 0.28
    if "lines" in data:
        text(
            slide, x + 0.14, body_y, width - 0.28,
            0.6 if "table" in data else body_height,
            data["lines"],
            size=11 if height < 2 else (12 if compact or width < 3.2 else 14),
            spacing=2 if height < 2 else (4 if compact else 6),
        )
    if "table" in data:
        offset = 0.68 if "lines" in data else 0
        table(
            slide, data["table"], x + 0.12, body_y + offset,
            width - 0.24, body_height - offset,
        )


def card_row(slide, cards, columns, y, height, compact=False):
    width = (SLIDE_WIDTH - 2 * MARGIN - GAP * (columns - 1)) / columns
    for index, item in enumerate(cards):
        row, column = divmod(index, columns)
        card(
            slide, item, MARGIN + column * (width + GAP),
            y + row * (height + GAP), width, height, compact,
        )


def add_slide(presentation, data):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color("background")
    box(slide, 0, 0, SLIDE_WIDTH, 0.12, "navy")
    text(
        slide, MARGIN, 0.32, 15.1, 0.55, [data["title"].upper()],
        size=25, ink="navy", bold=True,
    )
    text(
        slide, MARGIN, 0.94, 15.1, 0.42,
        [data["subtitle"] + "  |  " + PERIOD], size=13, ink="muted",
    )
    metrics = data["metrics"]
    width = (SLIDE_WIDTH - 2 * MARGIN - GAP * (len(metrics) - 1)) / len(metrics)
    accents = ["green", "blue", "orange", "purple", "light_green", "navy"]
    for index, (value, label) in enumerate(metrics):
        x = MARGIN + index * (width + GAP)
        box(slide, x, 1.55, width, 1.2, "white", "border")
        box(slide, x, 1.55, 0.06, 1.2, accents[index % len(accents)])
        text(slide, x + 0.15, 1.65, width - 0.3, 0.47, [value],
             size=26, ink="navy", bold=True)
        text(slide, x + 0.15, 2.16, width - 0.3, 0.55, [label],
             size=11, ink="muted")
    comparison = data.get("comparison")
    card_row(
        slide, data["cards"], data["columns"], 3.0,
        3.08 if comparison else 4.85, compact=bool(comparison),
    )
    if comparison:
        card_row(slide, comparison, 2, 6.26, 1.59, compact=True)
    box(slide, 0, 8.08, SLIDE_WIDTH, 0.92, "navy")
    text(
        slide, MARGIN, 8.18, 15.1, 0.42, [data["footer"]],
        size=12, ink="white", bold=True, align=PP_ALIGN.CENTER,
    )
    if data.get("footer_detail"):
        text(
            slide, MARGIN, 8.66, 15.1, 0.26, [data["footer_detail"]],
            size=10, ink="yellow", align=PP_ALIGN.CENTER,
        )
    return slide


def generate(output_file=OUTPUT_FILE):
    presentation = Presentation()
    presentation.slide_width = Inches(SLIDE_WIDTH)
    presentation.slide_height = Inches(SLIDE_HEIGHT)
    presentation.core_properties.title = "Business Review Presentation"
    presentation.core_properties.subject = "Editable monthly business review"
    for data in SLIDES:
        add_slide(presentation, data)
    presentation.save(output_file)
    return Path(output_file)


if __name__ == "__main__":
    print(f"Generated: {generate()}")
