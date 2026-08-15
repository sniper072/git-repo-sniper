#!/usr/bin/env python3
"""
Rebuild the Alfa-Bank certificate PDF with:
  - All dates replaced by 19 June 2026
  - Outgoing number changed from 9987-C/47540 to 2345-B/61823
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors

OUTPUT = "/workspace/certificate_modified.pdf"

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DejaVu", FONT_DIR + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", FONT_DIR + "DejaVuSans-Bold.ttf"))

W, H = A4
MARGIN_LEFT = 20 * mm
MARGIN_RIGHT = 20 * mm
MARGIN_TOP = 15 * mm
MARGIN_BOTTOM = 20 * mm


def style(name, font="DejaVu", size=9, leading=13, align=TA_LEFT,
          space_before=0, space_after=2, **kw):
    return ParagraphStyle(
        name,
        fontName=font,
        fontSize=size,
        leading=leading,
        alignment=align,
        spaceAfter=space_after,
        spaceBefore=space_before,
        **kw,
    )


s_normal   = style("normal")
s_bold     = style("bold",   font="DejaVuBold")
s_small    = style("small",  size=8, leading=11)
s_small_c  = style("smallc", size=8, leading=11, align=TA_CENTER)
s_right    = style("right",  align=TA_RIGHT)
s_center   = style("center", align=TA_CENTER, size=10, font="DejaVuBold", leading=14)
s_body     = style("body",   align=TA_JUSTIFY, size=9, leading=13)
s_footnote = style("foot",   size=7.5, leading=11)

def p(text, st=s_normal):
    return Paragraph(text, st)

def sp(h=4):
    return Spacer(1, h * mm)


def build():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=A4,
        leftMargin=MARGIN_LEFT,
        rightMargin=MARGIN_RIGHT,
        topMargin=MARGIN_TOP,
        bottomMargin=MARGIN_BOTTOM,
    )

    story = []

    # ── Top two-column block ─────────────────────────────────────────────────
    # Left column: bank details
    left_lines = [
        p("АО «АЛЬФА-БАНК»", s_bold),
        p("ИНН 7728168971", s_normal),
        p("ОГРН 1027700067328", s_normal),
        p("Генеральная лицензия Банка России", s_normal),
        p("№ 1326 от 16 января 2015 г.", s_normal),
    ]

    # Right column: outgoing number + date  ← CHANGED number & date
    right_lines = [
        p("Исх. № 2345-B/61823", s_right),          # ← number changed
        p("от «19» июня 2026 г.", s_right),          # ← date changed
    ]

    # Pack each column into a tiny nested table so line-breaks are preserved
    left_cell  = [line for line in left_lines]
    right_cell = [line for line in right_lines]

    col_w = (W - MARGIN_LEFT - MARGIN_RIGHT)
    half  = col_w / 2.0

    header_table = Table(
        [[left_cell, right_cell]],
        colWidths=[half, half],
        hAlign="LEFT",
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(sp(6))

    # ── Title ────────────────────────────────────────────────────────────────
    story.append(p("СПРАВКА", s_center))
    story.append(sp(1))
    story.append(
        p(
            "о наличии / отсутствии распоряжений по счетам, "
            "помещенных Банком в очереди распоряжений\u00b9",
            style("title_sub", align=TA_CENTER, size=9, leading=13),
        )
    )
    story.append(sp(5))

    # ── Body ─────────────────────────────────────────────────────────────────
    story.append(
        p(
            "Подтверждаем, что у ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "
            "\u201cРК-СИСТЕМС\u201d (ИНН 4501097640) в АО \u201cАЛЬФА-БАНК\u201d "
            "(далее \u2013 Банк) открыт(ы) счет(а):",
            s_body,
        )
    )
    story.append(sp(2))

    story.append(
        p("Расчетный счет (Российский рубль) № 40702810038290008548", s_normal)
    )
    story.append(sp(4))

    # ← Date changed here as well
    story.append(
        p(
            "При обслуживании Банком данного счета по состоянию на "
            "12:45:21 19.06.2026г.:",    # ← date changed
            s_body,
        )
    )
    story.append(sp(2))

    bullet_style = style(
        "bullet", size=9, leading=13, leftIndent=6 * mm, firstLineIndent=-3 * mm
    )
    story.append(
        p(
            "\u2013 распоряжения на списание денежных средств с данного счета, "
            "помещенные Банком в очередь распоряжений, ожидающих разрешения на "
            "проведение операций, отсутствует.",
            bullet_style,
        )
    )
    story.append(sp(2))
    story.append(
        p(
            "\u2013 распоряжения на списание денежных средств с данного счета, "
            "помещенные Банком в очередь неисполненных в срок распоряжений, "
            "отсутствуют",
            bullet_style,
        )
    )
    story.append(sp(10))

    # ── Signature block ──────────────────────────────────────────────────────
    sig_line = "_" * 45
    sig_table = Table(
        [[p(sig_line, s_normal), p("", s_normal)],
         [p("(Подпись сотрудника АО «АЛЬФА-БАНК»)", s_small_c), p("", s_normal)],
         [p("Белов А. Ю.", s_small_c), p("", s_normal)],
         [p("(Ф.И.О. сотрудника АО «АЛЬФА-БАНК»)", s_small_c), p("", s_normal)]],
        colWidths=[half, half],
        hAlign="LEFT",
    )
    sig_table.setStyle(TableStyle([
        ("ALIGN",        (0, 0), (0, -1), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 1),
    ]))
    story.append(sig_table)

    # ── Footnote rule + text ─────────────────────────────────────────────────
    story.append(sp(6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(sp(2))
    story.append(
        p(
            "\u00b9Помещение распоряжений в очереди осуществляется Банком в "
            "соответствии с Положением ЦБ РФ от 19 июня 2026 г. N 762-П "  # ← date changed
            "\u00abО правилах осуществления перевода денежных средств\u00bb",
            s_footnote,
        )
    )

    doc.build(story)
    print(f"PDF saved to {OUTPUT}")


if __name__ == "__main__":
    build()
