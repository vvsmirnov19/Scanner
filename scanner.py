import re

import openpyxl
from pdf2image import convert_from_path
from PIL import Image
import pytesseract

poppler_path = r'C:\\Program Files\\poppler\\Library\bin'

def convert_to_images(input_file, output_file):
    images = convert_from_path(input_file, poppler_path=poppler_path)
    image = images[0]
    image.save(output_file, "PNG")

def teseract_recognition(path_img):
    return pytesseract.image_to_string(Image.open(path_img), lang='rus+eng', config=r'--oem 3 --psm 6')


def save_text(text, name):
    with open(f'{name}.md', 'w', encoding='utf-8') as file:
        file.write(text)
    return

def extract_data(file_name):
    with open(file_name, 'r', encoding='utf-8') as file:
        output = list()
        strings = list()
        for line in file.readlines():
            strings.append(line.rstrip())
        format_string = ' '.join(strings)
        format_strings = format_string.split('[')
        for i in range(1, len(format_strings)):
            if '|' not in format_strings[i]:
                format_strings[i-1] = ' '.join([format_strings[i-1], format_strings[i]])
                del format_strings[i]
        for line in format_strings:
            tasks = dict()
            if not re.findall(r'(?<![\"=\w])(?:[^\W_]+)(?![\"=\w])', line):
                continue
            elif 'ЧАСТЬ' in line:
                part = line.split(' ')[1]
            else:
                tasks['Часть'] = part
                tasks['Номер вопроса'] = line.split('|')[0]
                tasks['Вопрос'] = line.split('|')[1]
                output.append(tasks)
        return output




def create_table_tasks(dictionary, path_name):
    path = path_name + '.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Лист 1'

    header = list(dictionary[0].keys())
    ws.append(header)
    for row in dictionary:
        ws.append([row[col] for col in header])

    wb.save(path)
    return path


def main():
    path_name = r'C:\\Dev\\scanner\\92e27331-e7eb-4794-947a-7fe3d2df18cd (1)'
    path_pdf = path_name + '.pdf'
    path_img = path_name + '.png'
    convert_to_images(path_pdf, path_img)
    name = r'92e27331-e7eb-4794-947a-7fe3d2df18cd (1)'
    save_text(teseract_recognition(path_img), name)
    md = path_name + '.md'
    create_table_tasks(extract_data(md), path_name)


if __name__ == "__main__":
    main()