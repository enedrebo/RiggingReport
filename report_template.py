from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_LEFT

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import os

# ---------- USER FONT ----------
FONT_NAME = "Helvetica"
LOGO_FILE = "logo_shield.png"

if os.path.isfile(f"{FONT_NAME}.ttf"):
    pdfmetrics.registerFont(TTFont(FONT_NAME, f"{FONT_NAME}.ttf"))

# ---------- Colors ----------
TEXT_COLOR = colors.Color(48/255, 51/255, 89/255)
HEADER_COLOR = colors.Color(204/255, 196/255, 176/255)
LINE_COLOR = colors.Color(255/255, 0/255, 228/255)

# ---------- Page setup ----------
PAGE_WIDTH, PAGE_HEIGHT = A4
HEADER_HEIGHT = PAGE_HEIGHT * 0.08
IMAGE_HEIGHT = PAGE_HEIGHT * 0.25

styles = getSampleStyleSheet()

# ---------- Styles ----------
title_style = ParagraphStyle(
    name="Title",
    parent=styles["Title"],
    fontName=FONT_NAME,
    textColor=TEXT_COLOR,
    alignment=TA_LEFT
)

subheader_style = ParagraphStyle(
    name="SubHeader",
    fontSize=10,
    spaceAfter=4,
    fontName=FONT_NAME,
    textColor=TEXT_COLOR
)

text_style = ParagraphStyle(
    name="Text",
    fontSize=9,
    fontName=FONT_NAME,
    textColor=TEXT_COLOR
)

formula_style = ParagraphStyle(
    name="Formula",
    fontSize=9,
    fontName=FONT_NAME,
    textColor=TEXT_COLOR
)

header_left_style = ParagraphStyle(
    name="HeaderLeft",
    fontSize=9,
    fontName=FONT_NAME,
    textColor=TEXT_COLOR,
    leading=11
)

header_center_style = ParagraphStyle(
    name="HeaderCenter",
    fontSize=11,
    fontName=FONT_NAME,
    textColor=TEXT_COLOR,
    alignment=1
)

# ---------- Header ----------
def _make_draw_header(project, date, author, approver):
    def draw_header(canvas, doc):
        width, height = A4

        canvas.setFillColor(HEADER_COLOR)
        canvas.rect(0, height - HEADER_HEIGHT, width, HEADER_HEIGHT, fill=1, stroke=0)

        left = Paragraph(
            f"Date: {date}<br/>Author: {author}<br/>Approver: {approver}",
            header_left_style
        )

        center = Paragraph(
            f"<b>{project}</b>",
            header_center_style
        )

        try:
            logo = Image(LOGO_FILE)
            logo._restrictSize(PAGE_WIDTH * 0.2, HEADER_HEIGHT * 0.8)
        except Exception:
            logo = Paragraph("", header_left_style)

        header_table = Table(
            [[left, center, logo]],
            colWidths=[PAGE_WIDTH*0.3, PAGE_WIDTH*0.4, PAGE_WIDTH*0.3]
        )

        header_table.setStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (2,0), (2,0), "RIGHT"),
        ])

        w, h = header_table.wrap(PAGE_WIDTH, HEADER_HEIGHT)
        header_table.drawOn(canvas, 0, height - HEADER_HEIGHT + (HEADER_HEIGHT - h)/2)

    return draw_header

# ---------- Layout ----------
TABLE_WIDTHS = [45*mm, 85*mm, 45*mm]
TOTAL_WIDTH = sum(TABLE_WIDTHS)

def create_table(rows):
    return Table(rows, colWidths=TABLE_WIDTHS, style=[
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("ALIGN", (2,0), (2,-1), "RIGHT"),
    ])

def subheader(text):
    return Table(
        [[Paragraph(f"<b>{text}</b>", subheader_style)]],
        colWidths=[TOTAL_WIDTH],
        style=[("LEFTPADDING", (0,0), (-1,-1), 0)]
    )

def make_row(r):
    return [
        Paragraph(r["desc"], text_style),
        Paragraph(r["formula"], formula_style),
        Paragraph(r["comment"], text_style),
    ]

# ---------- NEW: Page comment ----------
def page_comment(text):
    return Table(
        [[Paragraph(f"<i>{text}</i>", text_style)]],
        colWidths=[TOTAL_WIDTH],
        style=[
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 6),
        ]
    )

# ---------- Page builders ----------
def _add_image(elements, data, image_bytes, use_fallback=True):
    """Add image to elements.
    image_bytes (bytes) takes priority. If None and use_fallback is True,
    falls back to the filename in data['image'] (CLI mode only).
    """
    if image_bytes is not None:
        src = io.BytesIO(image_bytes)
    elif use_fallback:
        src = data.get("image")
    else:
        return
    if not src:
        return
    try:
        img = Image(src)
        img._restrictSize(TOTAL_WIDTH, IMAGE_HEIGHT)
        img.hAlign = "CENTER"
        elements.append(img)
        elements.append(Spacer(1, 12))
    except Exception:
        pass


def build_first_page(elements, data, image_bytes=None, use_fallback=True):
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"<b>{data['title']}</b>", title_style))

    line = Table([[""]], colWidths=[TOTAL_WIDTH], rowHeights=[1])
    line.setStyle([("BACKGROUND", (0,0), (-1,-1), LINE_COLOR)])
    elements.append(Spacer(1, 4))
    elements.append(line)
    elements.append(Spacer(1, 12))

    _add_image(elements, data, image_bytes, use_fallback)

    for block in data["sections"]:
        elements.append(subheader(block["title"]))
        elements.append(Spacer(1, 4))
        elements.append(create_table([make_row(r) for r in block["rows"]]))
        elements.append(Spacer(1, 12))

    if "comment" in data and data["comment"]:
        elements.append(Spacer(1, 10))
        elements.append(page_comment(data["comment"]))


def build_item_page(elements, data, image_bytes=None, use_fallback=True):
    elements.append(Spacer(1, 20))

    elements.append(Paragraph(f"<b>{data['title_line1']}</b>", subheader_style))
    elements.append(Paragraph(data["title_line2"], text_style))

    line = Table([[""]], colWidths=[TOTAL_WIDTH], rowHeights=[1])
    line.setStyle([("BACKGROUND", (0,0), (-1,-1), LINE_COLOR)])
    elements.append(Spacer(1, 4))
    elements.append(line)
    elements.append(Spacer(1, 12))

    _add_image(elements, data, image_bytes, use_fallback)

    for block in data["sections"]:
        elements.append(subheader(block["title"]))
        elements.append(Spacer(1, 4))
        elements.append(create_table([make_row(r) for r in block["rows"]]))
        elements.append(Spacer(1, 12))

    if "comment" in data and data["comment"]:
        elements.append(Spacer(1, 10))
        elements.append(page_comment(data["comment"]))


# ---------- Summary page ----------
SUMMARY_COL_WIDTHS = [30*mm, 60*mm, 65*mm, 20*mm]

def build_summary_page(elements, items_data):
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("<b>Summary table</b>", title_style))

    line = Table([[""]], colWidths=[TOTAL_WIDTH], rowHeights=[1])
    line.setStyle([("BACKGROUND", (0,0), (-1,-1), LINE_COLOR)])
    elements.append(Spacer(1, 4))
    elements.append(line)
    elements.append(Spacer(1, 12))

    header_row = [
        Paragraph("<b>Item number</b>", subheader_style),
        Paragraph("<b>Item description</b>", subheader_style),
        Paragraph("<b>Specification</b>", subheader_style),
        Paragraph("<b>Utilization</b>", subheader_style),
    ]
    data_rows = [header_row]
    for item in items_data:
        uf_val = item.get("uf")
        uf_str = f"{uf_val:.5g}" if uf_val is not None else "—"
        data_rows.append([
            Paragraph(item.get("item_no", ""), text_style),
            Paragraph(item.get("item_desc", ""), text_style),
            Paragraph(item.get("det_desc", ""), text_style),
            Paragraph(uf_str, text_style),
        ])

    tbl = Table(data_rows, colWidths=SUMMARY_COL_WIDTHS, style=[
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("BACKGROUND", (0,0), (-1,0), HEADER_COLOR),
        ("LINEBELOW", (0,0), (-1,0), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.Color(0.95, 0.95, 0.95)]),
    ])
    elements.append(tbl)


# ---------- BUILD ----------
def build_pdf(loads_data, items_data, output, page_images=None):
    """
    output     — file path string (CLI) or BytesIO (Streamlit).
    page_images — list of image bytes or None, one entry per page
                  (index 0 = loads page, 1..n = item pages).
    """
    # use_fallback=True in CLI mode (page_images not provided) so images are
    # loaded from disk by filename. False in Streamlit mode so only explicitly
    # uploaded images are used.
    use_fallback = page_images is None
    if page_images is None:
        page_images = []

    def _img(i):
        return page_images[i] if i < len(page_images) else None

    doc = SimpleDocTemplate(
        output,
        leftMargin=40,
        rightMargin=40,
        topMargin=HEADER_HEIGHT + 10,
        bottomMargin=40
    )

    draw_header = _make_draw_header(
        loads_data["project"],
        loads_data["date"],
        loads_data["author"],
        loads_data["approver"],
    )

    elements = []
    build_first_page(elements, loads_data, _img(0), use_fallback)

    for i, item in enumerate(items_data):
        elements.append(PageBreak())
        build_item_page(elements, item, _img(i + 1), use_fallback)

    if items_data:
        elements.append(PageBreak())
        build_summary_page(elements, items_data)

    doc.build(elements, onFirstPage=draw_header, onLaterPages=draw_header)