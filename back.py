# -*- coding: utf-8 -*-

from docx import Document
import pandas as pd
import re

doc = Document(r'C:\Users\asus\Documents\NDT YKR\NDT\REPORTS 2021\04-YKR-ON-UTT-21-788 (KUT 560) 8.12ap.docx')

all_tables = doc.tables

# создаем пустой словарь под данные таблиц
data_tables = {i: None for i in range(len(all_tables))}

for i, table in enumerate(all_tables):
    # создаем список строк для таблицы `i` (пока пустые)
    data_tables[i] = [[] for _ in range(len(table.rows))]
    # проходимся по строкам таблицы `i`
    for j, row in enumerate(table.rows):
        for cell in row.cells:
            data_tables[i][j].append(cell.text)

# снимаем ограничения на вывод таблицы
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

# Словарь очищенных таблиц с данными
clear_tables = {}
# Число таблиц с очищенными данными
number_clear_tables = 0

# проходим по всем найденным таблицам (спискам списков)
for i in data_tables:
    # проходим по всем  спискам
    for ii in data_tables[i]:
        # проходим по всем элементам списка
        for iii in ii:
            # выбираем только таблицы с нужными данными
            if re.match(r'Nominal thickness', iii) or re.match(r'DIA', iii):
                clear_tables[number_clear_tables] = data_tables[i]
                number_clear_tables += 1


# количество очищенных таблиц с данными
len_tables = len(clear_tables)
# очистка таблиц
for i in range(len(clear_tables)):
    for ii in clear_tables[i]:
        # очистка от общей первой строки
        if len(set(clear_tables[i][0])) == 1:
            clear_tables[i].pop(0)

# формируем таблицы
# for i in range(len(clear_tables)):
#     s = pd.DataFrame(clear_tables[i])
#     # s = s.drop(index=[0])
#     print(s)

