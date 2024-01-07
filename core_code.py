from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os


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

def extract_tables(source_filenames, keyword, output_filename):
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

    new_doc.save(output_filename)


def select_files():
    """
    Открывает диалоговое окно для выбора файлов и сохраняет выбранные пути.
    """
    filetypes = [('Word files', '*.docx'), ('All files', '*.*')]
    filenames = filedialog.askopenfilenames(title='Open files', initialdir='/', filetypes=filetypes)
    return list(filenames)

def start_extraction():
    source_filenames = select_files()
    if not source_filenames:
        messagebox.showwarning("Warning", "No files selected!")
        return

    keyword = simpledialog.askstring("Input", "Enter the keyword for searching tables", parent=root)
    if not keyword:
        messagebox.showwarning("Warning", "Keyword not entered!")
        return

    output_filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if not output_filename:
        messagebox.showwarning("Warning", "No save location selected!")
        return

    try:
        extract_tables(source_filenames, keyword, output_filename)
        messagebox.showinfo("Success", f"Tables extracted successfully and saved to {output_filename}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI setup
root = tk.Tk()
root.title("Word Table Extractor")

start_button = tk.Button(root, text="Start Extraction", command=start_extraction)
start_button.pack()

root.mainloop()
