import locale
from collections import namedtuple

import openpyxl as op

from date import Date

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

Info = namedtuple('Info', ['students', 'date_start', 'date_end'])


def get_students(sheet) -> list:
    """Получение списка студентов"""
    table_index = 4
    students = []
    while sheet[f'A{table_index}'].value is not None:
        a = sheet[f'A{table_index}'].value
        table_index += 1
        students.append(a)
    if len(students) == 0:
        print('Список студентов пустой. Заполните файл Ученики')
        exit()
    elif len(students) == 1:
        students.append('')
    return students


def get_dates(sheet) -> list:
    """Получение даты начала и окончания обучения"""
    try:
        start_date = Date(date_time=sheet['B1'].value)
        end_date = Date(date_time=sheet['C1'].value)
        return [start_date, end_date]
    except AttributeError:
        print('Данные введены не в формате даты')
        exit()


def read_information_document() -> Info:
    """Чтение списка студентов и даты из файла Ученики.xlsx"""
    file_path = 'Ученики.xlsx'
    # TODO: Проверить, что файл существует
    try:
        excel_doc = op.open(filename=file_path, data_only=True)
        sheet = excel_doc[excel_doc.sheetnames[0]]

        data = Info(get_students(sheet), *get_dates(sheet))

        return data
    except FileNotFoundError:
        print('Не существует файл "Ученики.xlsx"')
        exit()


def main():
    wb = op.Workbook()
    ws = wb.active
    ws.title = 'Дневник наблюдений'

    header = ['Дата', 'Имя ученика', 'Эмоциональное состояние ребёнка',
              'Деятельность, затруднения, достижения',
              ]
    ws.append(header)

    data = read_information_document()

    start_date = data.date_start
    end_date = data.date_end
    current_date = start_date

    while current_date < end_date:
        ws.append([current_date.get_datetime_with_format(), data.students[0]])
        ws.append([current_date.get_weekday(), data.students[1]])
        for i in range(2, len(data.students)):
            ws.append(['', data.students[i]])
        current_date = current_date.next_workday()

    wb.save('Дневник.xlsx')


if __name__ == '__main__':
    main()
