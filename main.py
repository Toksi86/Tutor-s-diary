import openpyxl as op

from date import Date


def get_students(sheet):
    i = 2
    students = []
    while sheet[f'A{i}'].value is not None:
        a = sheet[f'A{i}'].value
        i += 1
        students.append(a)
    if len(students) == 0:
        print('Список студентов пустой. Заполните файл Ученики')
    elif len(students) == 2:
        students.append('')
    return students


def date_processing(date: str):
    data_list = [date.day, date.month, date.year]
    return data_list


def get_dates(sheet):
    start_date = date_processing(sheet['B2'].value)
    end_date = date_processing(sheet['C2'].value)
    return [start_date, end_date]


def read_information_document():
    file_path = 'Ученики.xlsx'
    excel_doc = op.open(filename=file_path, data_only=True)
    sheetnames = excel_doc.sheetnames
    sheet = excel_doc[sheetnames[0]]

    students = get_students(sheet)

    dates = get_dates(sheet)

    datas = {
        'students': students,
        'start': dates[0],
        'end': dates[1],
    }

    return datas


def main():
    wb = op.Workbook()
    ws = wb.active
    ws.title = 'Дневник наблюдений'

    header = ['Дата', 'Имя ученика', 'Эмоциональное состояние ребёнка',
              'Деятельность, затруднения, достижения',
              ]
    ws.append(header)

    datas = read_information_document()

    start_date = Date(datas['start'][2], datas['start'][1], datas['start'][0])
    end_date = Date(datas['end'][2], datas['end'][1], datas['end'][0])
    current_date = start_date

    while current_date < end_date:
        ws.append([current_date.get_datetime_with_format(), datas['students'][0]])
        ws.append([current_date.get_weekday(), datas['students'][1]])
        for i in range(2, len(datas['students'])):
            ws.append(['', datas['students'][i]])
        current_date = current_date.next_workday()

    wb.save('Дневник.xlsx')


if __name__ == '__main__':
    main()
