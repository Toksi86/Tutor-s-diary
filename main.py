from openpyxl import Workbook

from date import Date

wb = Workbook()
ws = wb.active
ws.title = 'Дневник наблюдений'


header = ['Дата', 'Имя ученика', 'Эмоциональное состояние ребёнка',
          'Деятельность, затруднения, достижения',
          ]

while True:
    number_of_students = input('Введите число обучающихся: ')
    if number_of_students.isdigit():
        number_of_students = int(number_of_students)
        break


# TODO: Получать учеников из другого excel файла
students = []

for _ in range(number_of_students):
    student = input('Введите ФИО студента: ')
    students.append(student)

while len(students) < 2:
    students.append('')

ws.append(header)

# TODO: дать возможность пользователю ввести дату начала и конца обучения из другого файла excel
current_date = Date(2023, 9, 4)

while current_date < Date(2024, 5, 24):
    ws.append([current_date.get_datetime_with_format(), students[0]])
    ws.append([current_date.get_weekday(), students[1]])
    for i in range(2, len(students)):
        ws.append(['', students[i]])
    current_date = current_date.next_workday()

wb.save('Дневник.xlsx')
