from docx import Document
from docx.oxml.ns import qn

doc = Document('templates/Word_template_copy.docx')
body = doc.element.body

tbl_count = 0
for idx, child in enumerate(body):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag == 'tbl':
        rows = child.findall(qn('w:tr'))
        first_row_cells = []
        if rows:
            for tc in rows[0].findall(qn('w:tc')):
                text = ''.join(p.text or '' for p in tc.iter() if p.text)
                first_row_cells.append(text[:50])
        print(f'Body idx {idx}: TABLE {tbl_count} ({len(rows)} rows) First row: {first_row_cells}')
        tbl_count += 1
    elif tag == 'p':
        text = ''.join(t.text or '' for t in child.iter() if t.text)
        if text.strip():
            print(f'Body idx {idx}: PARA "{text[:100]}"')
