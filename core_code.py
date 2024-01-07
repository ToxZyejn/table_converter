from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn



def set_cell_border(cell, **kwargs):
    """
    Установка границ ячейки таблицы
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Создаем элемент 'w:tcBorders' для границ
    tcBorders = OxmlElement('w:tcBorders')
    for key, value in kwargs.items():
        tag = 'w:{}'.format(key)
        element = OxmlElement(tag)
        element.set(qn('w:val'), value)
        tcBorders.append(element)

    tcPr.append(tcBorders)

def apply_borders_to_table(table):
    """
    Применяем границы ко всем ячейкам в таблице
    """
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top="single", bottom="single", start="single", end="single")

def extract_tables(source_filenames, keyword):
    new_doc = Document()
    for filename in source_filenames:
        doc = Document(filename)
        for table in doc.tables:
            if keyword in [cell.text for row in table.rows for cell in row.cells]:
                new_tbl = new_doc.add_table(rows=0, cols=len(table.columns))
                for row in table.rows:
                    cells = row.cells
                    new_row = new_tbl.add_row().cells
                    for idx, cell in enumerate(cells):
                        new_row[idx].text = cell.text
                apply_borders_to_table(new_tbl)
                new_doc.add_paragraph()

    new_doc.save('extracted_tables.docx')


# Список файлов для обработки
source_filenames = ['your_word_file_1.docx', ..., 'your_word_file_10.docx']

# Ключевое слово для поиска таблиц
keyword = 'keyword_for_searching_table'

# Вызов функции
extract_tables(source_filenames, keyword)
