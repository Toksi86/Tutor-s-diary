# Дневник тьютора
## _Программа позволяющая создать дневник тьютора в excel для тьюторов, работающих с детьми с ограниченными возможностями здоровья_

Дневник тьютора — это excel таблица, которая содержит четыре столбца:
- Дата
- Имя ученика
- Эмоциональное здоровье ребёнка
- Деятельность, затруднения, достижения

На каждый будний день создаётся набор данных, состоящих из даты и дня недели, перечня всех детей и возможности
оставить комментарий для каждого ребёнка о его состоянии и деятельности.

| Дата        | Имя ученика  | Эмоциональное состояние ребёнка |    Деятельность, затруднения, достижения |
|-------------|:------------:|--------------------------------:|-----------------------------------------:|
| 04-09-2023  | Вася Петров  |                         Хорошее |                           Нарисовал маму |
| понедельник | Петя Иванов  |                        Отличное |       Решил сложную задачу по математике |
|             | Коля Сидоров |              Удовлетворительное |         Помог учительнице прибрать класс |


## Установка
Для корректной работы необходимо установить модуль openpyxl

```sh
pip install openpyxl
```


## Запуск

- Заполните файл "Ученики.xlsx": в первом столбце ФИО учеников, во втором и третьем дата начала и завершения 
обучения соответственно. 
- Запустите программу main.py

По завершению выполнения в рабочей директории создастся файл Дневник.xlsx

