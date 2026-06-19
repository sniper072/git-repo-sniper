"""
Update Tochka Bank certificate PDF.

Changes:
- Removes the "есть действующие ограничения" block and all restriction details
- Replaces body with: no active restrictions as of 19.06.2026
- Клиент line and all header/footer content are preserved unchanged
"""
import fitz

SRC = 'Spravka_Tochka_free_a445.pdf'
DST = 'Spravka_Tochka_free_updated.pdf'
FONT_FILE = '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf'

NEW_BODY = (
    'В Обществе с ограниченной ответственностью «Банк Точка» по состоянию на '
    '19.06.2026 г. на счёте 40702810802500044569 RUB отсутствуют действующие '
    'ограничения на распоряжение деньгами'
)

doc = fitz.open(SRC)
page = doc[0]

# Redact: original "В Обществе... деньгами:" span block
page.add_redact_annot(fitz.Rect(55, 230, 570, 278), fill=(1, 1, 1))
# Redact: entire restriction details block (tax hold, amounts, UIDs)
page.add_redact_annot(fitz.Rect(55, 278, 550, 395), fill=(1, 1, 1))
page.apply_redactions()

# Insert replacement body text with proper Cyrillic/Unicode support
font = fitz.Font(fontfile=FONT_FILE)
tw = fitz.TextWriter(page.rect)
overflow = tw.fill_textbox(
    fitz.Rect(60, 233.6, 565, 400),
    NEW_BODY,
    font=font,
    fontsize=12,
    align=0,
)
if overflow:
    print(f"WARNING: {len(overflow)} chars did not fit!")
tw.write_text(page, color=(0, 0, 0))

doc.save(DST, garbage=4, deflate=True)
doc.close()
print(f"Saved: {DST}")
