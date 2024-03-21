# Макрос по разбиению pdf-файла на отдельные файлы (листы) и их переименованию (в соответствии с артикулом)
# Читаем PDF-файл, разбиваем его на отдельные файлы
# Читаем каждый отдельный файл, ищем там артикул и переименовываем файл по артикулу
# Если артикул в файле не найден, то файл остаётся с прежним названием
# Файл шаблона кидать в папку (строка 13)

import os
import re
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import getpass

os.chdir(rf'C:\Users\{getpass.getuser()}\Desktop\DS_LOCAL\DS_BS')

# Разбиваем файл с даташитами на отдельные файлы
for file in os.listdir():
    if file.startswith('DS') and file.endswith('.pdf'):
        input_PDF = PdfReader(open(file, 'rb'))

        for i in range(len(input_PDF.pages)):
            output = PdfWriter()
            new_File_PDF = input_PDF.pages[i]
            output.add_page(new_File_PDF)
            output_Name_File = "pdffile_" + str(i + 1) + ".pdf"
            with open(output_Name_File, 'wb') as outputStream:
                output.write(outputStream)

# Переименовываем файлы
for file in os.listdir():
    if file.endswith('.pdf') and not file.startswith('DS'):
        print(file)
        with pdfplumber.open(file) as pdf_file:
            text = pdf_file.pages[0].extract_text()
            if re.search(r'C3.\w{2,3}\d{3,4}', text):
                print(re.search(r'C3.\w{2,3}\d{3,4}', text)[0])
                article = re.search(r'C3.\w{2,3}\d{3,4}', text)[0]

        os.rename(file, 'DS_' + article + '.pdf')

