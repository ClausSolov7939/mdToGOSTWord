from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Mm
import json

fin = open("../input/document.md", "r")
styleFile = open("style.json", encoding='utf-8')

document = Document()
styles = document.styles

styleData = json.load(styleFile)

for newStyle in styleData.keys():
    currentStyle = styles.add_style(newStyle, WD_STYLE_TYPE.PARAGRAPH)
    print(currentStyle)
    # -------- paragraph setting -----------
    font_format = currentStyle.font
    font_style = styleData[newStyle]["font"]
    font_format.name = font_style["name"]
    font_format.size = Pt(font_style["size"])
    font_format.bold = font_style["bold"]
    font_format.italic = font_style["italic"]
    font_format.underline = font_style["underline"]
    font_format.all_caps = font_style["all_caps"]
    # -------- paragraph setting -----------
    paragraph_format = currentStyle.paragraph_format
    paragraph_style = styleData[newStyle]["paragraph"]
    # "alignment": "Left",

    paragraph_format.first_line_indent = Mm(paragraph_style["first_line_indent"])
    paragraph_format.left_indent = Mm(paragraph_style["left_indent"])
    paragraph_format.right_indent = Mm(paragraph_style["right_indent"])
    paragraph_format.space_before = Mm(paragraph_style["space_before"])
    paragraph_format.space_after = Mm(paragraph_style["space_after"])
    paragraph_format.line_spacing = Mm(paragraph_style["line_spacing"])
    # Запрет висящих строк
    currentStyle.widow_control = paragraph_style["widow_control"]
    # Не отрывать от следующего
    currentStyle.keep_with_next = paragraph_style["keep_with_next"]
    # Не разрывать абзац
    currentStyle.keep_together = paragraph_style["keep_together"]
    # с новой страницы
    currentStyle.page_break_before = paragraph_style["page_break_before"]
    print(paragraph_style["page_break_before"], currentStyle.page_break_before)


print(document.styles["Заголовок 1"])
print(document.styles["Заголовок 1"].paragraph_format.page_break_before)


for line in fin:
    splitedLine = line.split()
    print(splitedLine)
    if len(splitedLine) != 0:  # if line is not empty
        if "#" in splitedLine[0]:  # if line contains a Header
            header = document.add_paragraph(splitedLine[1:])
            header.style = document.styles[f'Заголовок {len(splitedLine[0])}']
# p = document.add_paragraph('A plain paragraph having some ')
# p.add_run('bold').bold = True
# p.add_run(' and some ')
# p.add_run('italic.').italic = True
#
# document.add_heading('Heading, level 1', level=1)
# document.add_paragraph('Intense quote', style='Intense Quote')
#
# document.add_paragraph(
#     'first item in unordered list', style='List Bullet'
# )
# document.add_paragraph(
#     'first item in ordered list', style='List Number'
# )
#
# # document.add_picture('monty-truth.png', width=Inches(1.25))
#
# records = (
#     (3, '101', 'Spam'),
#     (7, '422', 'Eggs'),
#     (4, '631', 'Spam, spam, eggs, and spam')
# )
#
# table = document.add_table(rows=1, cols=3)
# hdr_cells = table.rows[0].cells
# hdr_cells[0].text = 'Qty'
# hdr_cells[1].text = 'Id'
# hdr_cells[2].text = 'Desc'
# for qty, id, desc in records:
#     row_cells = table.add_row().cells
#     row_cells[0].text = str(qty)
#     row_cells[1].text = id
#     row_cells[2].text = desc
#
# document.add_page_break()

document.save('../output/demo.docx')
