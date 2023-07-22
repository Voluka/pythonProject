from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font, Color
from openpyxl.utils import get_column_letter

id = 9
number_of_id_row = int
name_of_column = int

formatting = []
numbering = []

def get_numbering_paragraphs(doc):
    numbering_paragraphs = []
    for paragraph in doc.paragraphs:
        if paragraph._p.pPr is None:
            continue
        numbering_paragraphs.append(paragraph)
    return numbering_paragraphs

def parse_word_file1(file_path):
    doc = Document(file_path)
    numbering_paragraphs = get_numbering_paragraphs(doc)
    for paragraph in numbering_paragraphs:
            numbering_text = paragraph.text
            # print(numbering_text)



def parse_word_file(word_file, excel_file, sheet_name, start_keyword, end_keyword):
    # Загружаем документ Word
    doc = Document(word_file)
    # Загружаем книгу Excel и выбираем нужный лист
    workbook = load_workbook(excel_file)
    sheet = workbook[sheet_name]
    print(sheet.max_column)
    for i in range(1, sheet.max_column):
        print(sheet.cell(row=1, column=i).value)
        if sheet.cell(row=1, column=i).value == 'показания к применению':
            name_of_column = i
    print(name_of_column)
    for i in range(1, sheet.max_row):
        if sheet.cell(row=i, column=1).value == id:
            number_of_id_row = i
    print(number_of_id_row)

    # Ищем нужные абзацы в документе Word
    start = False
    text = ''
    for paragraph in doc.paragraphs:
        if start_keyword in paragraph.text:
            start = True
        elif end_keyword in paragraph.text:

            start = False
            # Вставляем текст в следующую свободную ячейку Excel
            sheet.cell(row=number_of_id_row, column=name_of_column).value = text
            text = ''
        elif start:
            text += paragraph.text


    # Сохраняем изменения в файле Excel
    workbook.save(excel_file)
