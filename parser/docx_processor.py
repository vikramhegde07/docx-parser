from docx import Document
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.shared import Inches
import base64
from io import BytesIO
from PIL import Image
from lxml import etree

def iter_block_items(parent):
    for child in parent.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield DocxTable(child, parent)

def get_paragraph_alignment(paragraph):
    alignment = paragraph.alignment
    if alignment == 0:
        return "text-left"
    elif alignment == 1:
        return "text-center"
    elif alignment == 2:
        return "text-right"
    elif alignment == 3:
        return "text-justify"
    return "text-left"  # default


def get_inline_formatting(run):
    text = run.text
    if not text:
        return ''

    styles = []

    # Inline formatting
    if run.bold:
        text = f"<strong>{text}</strong>"
    if run.italic:
        text = f"<em>{text}</em>"
    if run.underline:
        text = f"<u>{text}</u>"

    # Font size (in half-points, so 24 => 12px)
    rPr = run._element.rPr
    if rPr is not None:
        sz = rPr.find(qn('w:sz'))
        if sz is not None:
            try:
                size = int(sz.get(qn('w:val')))
                px = int(size / 2)
                styles.append(f"font-size:{px}px")
            except Exception:
                pass

    style_str = f' style="{";".join(styles)}"' if styles else ''
    return f"<span{style_str}>{text}</span>"


def get_numbering_format(paragraph):
    pPr = paragraph._p.pPr
    if pPr is not None and pPr.numPr is not None:
        numPr = pPr.numPr
        numId = numPr.numId.val if numPr.numId is not None else None
        ilvl = numPr.ilvl.val if numPr.ilvl is not None else 0
        fmt = 'bullet' if numId and int(numId) % 2 == 0 else 'decimal'  # crude detection
        return fmt, ilvl
    return None, None

def process_paragraph(paragraph):
    return ''.join([get_inline_formatting(run) for run in paragraph.runs])

def process_table(table):
    html = ["<table class='docx-table'>"]
    for row in table.rows:
        html.append("<tr>")
        for cell in row.cells:
            cell_html = []
            for para in cell.paragraphs:
                text = process_paragraph(para).strip()
                if text:
                    style = para.style.name if para.style else ''
                    tag = 'p'
                    if style.startswith("Heading"):
                        level = ''.join(filter(str.isdigit, style)) or '1'
                        tag = f"h{level}"
                    cell_html.append(f"<{tag} class='docx--{style.replace(' ', '-').lower()}'>{text}</{tag}>")
            html.append(f"<td>{''.join(cell_html)}</td>")
        html.append("</tr>")
    html.append("</table>")
    return '\n'.join(html)

def extract_inline_images(paragraph):
    images = []
    try:
        # Search for all <w:drawing> tags inside the paragraph
        drawings = paragraph._element.xpath('.//w:drawing',)

        for drawing in drawings:
            # Look for <a:blip> inside the drawing
            blips = drawing.xpath('.//a:blip')
            for blip in blips:
                r_embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if not r_embed:
                    continue

                part = paragraph.part.related_parts.get(r_embed)
                if not part:
                    continue

                image_bytes = part.blob
                content_type = part.content_type.split('/')[-1]  # e.g., png, jpeg

                base64_image = base64.b64encode(image_bytes).decode('utf-8')
                src = f"data:image/{content_type};base64,{base64_image}"
                images.append(src)

    except Exception as e:
        print(f"[Image Parsing Error]: {e}")
    return images


def parse_docx(filepath, upload_dir="uploads"):
    doc = Document(filepath)
    html = []
    list_stack = []

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            style = block.style.name if block.style else ''
            alignment_class = get_paragraph_alignment(block)
            tag = 'p'

            list_type, level = get_numbering_format(block)

            text = process_paragraph(block).strip()
            inline_images = extract_inline_images(block)

            if list_type:
                while list_stack and list_stack[-1]['level'] > level:
                    html.append(f"</{list_stack[-1]['type']}>")
                    list_stack.pop()

                if not list_stack or list_stack[-1]['level'] < level or list_stack[-1]['type'] != ('ul' if list_type == 'bullet' else 'ol'):
                    tag_type = 'ul' if list_type == 'bullet' else 'ol'
                    html.append(f"<{tag_type}>")
                    list_stack.append({'type': tag_type, 'level': level})

                item_html = text + ''.join(f'<img src="{src}" />' for src in inline_images)
                html.append(f"<li>{item_html}</li>")

            else:
                while list_stack:
                    html.append(f"</{list_stack[-1]['type']}>")
                    list_stack.pop()

                if style.startswith("Heading"):
                    level = ''.join(filter(str.isdigit, style)) or '1'
                    tag = f"h{level}"

                style_class = f"docx--{style.replace(' ', '-').lower()}"
                content = text + ''.join(f'<img src="{src}" />' for src in inline_images)
                html.append(f"<{tag} class='{style_class} {alignment_class}'>{content}</{tag}>")

        elif isinstance(block, DocxTable):
            while list_stack:
                html.append(f"</{list_stack[-1]['type']}>")
                list_stack.pop()
            html.append(process_table(block))

    while list_stack:
        html.append(f"</{list_stack[-1]['type']}>")
        list_stack.pop()

    return '\n'.join(html)
