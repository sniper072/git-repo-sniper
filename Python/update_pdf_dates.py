"""
Update dates in AlfaBank certificate PDF.
Changes 10.02.2026 -> 19.06.2026 (document date and account status date).
Only dates are modified; all other content is left untouched.
"""
import fitz

SRC = 'AlfaBank_spravka_dcb3.pdf'
DST = 'AlfaBank_spravka_dcb3_updated.pdf'
FONT_FILE = '/usr/share/fonts/truetype/croscore/Tinos-Regular.ttf'

doc = fitz.open(SRC)
page = doc[0]

blocks = page.get_text('dict')
to_replace = []

for block in blocks['blocks']:
    if 'lines' in block:
        for line in block['lines']:
            for span in line['spans']:
                text = span['text']
                new_text = None
                if text == 'от «10» февраля 2026 г.':
                    new_text = 'от «19» июня 2026 г.'
                elif '10.02.2026' in text:
                    new_text = text.replace('10.02.2026', '19.06.2026')

                if new_text:
                    to_replace.append({
                        'bbox': fitz.Rect(span['bbox']),
                        'origin': span['origin'],
                        'text': new_text,
                        'size': span['size'],
                    })
                    print(f"OLD: {repr(text)}")
                    print(f"NEW: {repr(new_text)}\n")

# Redact old text areas (fill with white)
for r in to_replace:
    page.add_redact_annot(r['bbox'], fill=(1, 1, 1))
page.apply_redactions()

# Insert new text with proper Unicode/Cyrillic support
font = fitz.Font(fontfile=FONT_FILE)
tw = fitz.TextWriter(page.rect)
for r in to_replace:
    tw.append(r['origin'], r['text'], font=font, fontsize=r['size'])
tw.write_text(page, color=(0, 0, 0))

doc.save(DST, garbage=4, deflate=True)
doc.close()
print(f"Saved: {DST}")
