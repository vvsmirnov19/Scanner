import re

import openpyxl

def extract_data(file_name):
    with open(file_name, 'r', encoding='utf-8') as file:
        output = list()
        strings = list()
        for line in file.readlines():
            strings.append(line.rstrip())
        format_string = ' '.join(strings)
        parts = format_string.split('\\section*')
        for part in parts:
            part_id = re.findall(r'\{.*ЧАСТЬ\s(\d+).*\}', part)[0]
            tasks_str = re.split(r'\s([a-zA-Zа-яА-Я]\d)+\s', part)[1:]
            for i in range(1, len(tasks_str), 2):
                task_set = dict()
                image = re.findall(r'\!\[\]\((.+)\)', tasks_str[i])
                task_set['Часть'] = part_id
                task_set['Номер вопроса'] = tasks_str[i-1]
                task_set['Вопрос'] = re.split(r'\!\[\]\(.+\)', tasks_str[i])[0] if image else tasks_str[i]
                image = re.findall(r'\!\[\]\((.+)\)', tasks_str[i])
                task_set['Рисунок'] = image[0] if image else ' '
                output.append(task_set)
        return output


def create_table_tasks(dictionary, path_name):
    path = path_name[:-3] + '.xlsx'
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
    path_name = r'92e27331-e7eb-4794-947a-7fe3d2df18cd (1) (1).md'
    extract_data(path_name)
    create_table_tasks(extract_data(path_name), path_name)


if __name__ == "__main__":
    main()