# -*- coding: utf-8 -*-

from docx import Document
import re
import glob
import sqlite3


def add_table():
    # задаём папку для поиска репортов с расширением docx для Word 2013 и старше
    target_dir_docx = r'C:\Users\asus\Documents\NDT YKR\NDT UTT\**\*.docx'
    # target_dir_docx = r'C:\Users\asus\Documents\NDT YKR\NDT UTT\REPORTS 2020\*.docx'

    # присваиваем переменной список найденных файлов с расширением docx
    list_find_docx = glob.glob(target_dir_docx)

    # счётчик репортов с пустыми таблицами
    zero_table = 0
    # список репортов с пустыми таблицами
    list_zero_table = []
    # счётчик репортов с нарушением структуры
    distract_structure = 0
    # список репортов с нарушением структуры
    list_distract_structure = []
    # список ошибок в названиях столбцов таблиц
    message_column = []
    # список таблиц, которые не должны быть записаны
    dont_save_tables = []
    # список репортов с нарушением диапазона ячеек
    list_cells = []
    # сообщения о возникших проблемах и ошибках
    message_mistake = []

    # проверяем найденные репорты
    for unit_list_find_docx in list_find_docx:
        # работаем только если файл имеет допустимое имя, начинающееся с "04-YKR..."
        if re.findall(r'04-YKR', unit_list_find_docx):
            doc = Document(unit_list_find_docx)
            # Извлекаем номер и дату репорта из верхнего колонтитула
            head_paragraph = doc.sections[0].header.tables
            # создаем пустой словарь под данные верхнего колонтитула
            data_header = {i: None for i in range(0, len(head_paragraph))}
            for i, table in enumerate(head_paragraph):
                # создаем список строк для таблицы `i` (пока пустые)
                data_header[i] = [[] for _ in range(0, len(table.rows))]
                # проходимся по строкам таблицы `i`
                for j, row in enumerate(table.rows):
                    for cell in row.cells:
                        data_header[i][j].append(cell.text)
            pp = 0
            # словарь номера репорта, даты репорта, номера Work Order, номера линии, номера чертежа
            rep_number = {}
            for i in data_header[0]:
                p = 0
                for ii in i:
                    # если есть слово "Date" и любая цифра, то дата находится в этой же ячейке
                    if re.match(r'Date', ii) and re.findall(r'\d', ii):
                        rep_number['report_date'] = data_header[0][pp][p][-11:]
                    elif re.match(r'Date', ii) and re.findall(r'\D', ii):
                        rep_number['report_date'] = data_header[0][pp][p + 1]
                    # если есть слово "Report" и любая цифра, то номер репорта находится в этой же ячейке
                    if re.match(r'Report', ii) and re.findall(r'\d', ii):
                        rep_number['report_number'] = data_header[0][pp][p][11:]
                    elif re.match(r'Report', ii) and re.findall(r'\D', ii):
                        rep_number['report_number'] = '_' + data_header[0][pp][p + 1]
                        rep_number['report_number'] = re.sub('-', '_', rep_number['report_number'])
                    # если есть слово "Work" и любая цифра, то номер Work Order ордера находится в этой же ячейке
                    if re.match(r'Work', ii) and re.findall(r'\d', ii):
                        rep_number['work_order'] = data_header[0][pp][p][-8:]
                    elif re.match(r'Work', ii) and re.findall(r'\D', ii):
                        rep_number['work_order'] = data_header[0][pp][p + 1]
                    p += 1
                pp += 1

            # извлекаем необходимые данные из репорта
            # переменная со всеми таблицами в репорте
            all_tables = doc.tables
            # создаем пустой словарь под данные таблиц
            data_tables = {i: None for i in range(0, len(all_tables))}
            for i, table in enumerate(all_tables):
                # создаем список строк для таблицы `i` (пока пустые)
                data_tables[i] = [[] for _ in range(0, len(table.rows))]
                # проходимся по строкам таблицы `i`
                for j, row in enumerate(table.rows):
                    try:
                        for cell in row.cells:
                            data_tables[i][j].append(cell.text)
                    except IndexError:
                        mess = 'Ошибка в диапазоне ячеек репорта ' + rep_number['report_number']
                        message_mistake.append(mess)
                        list_cells.append(rep_number['report_number'])
                        dont_save_tables.append(rep_number['report_number'])
                        break

            # словарь таблиц с необходимыми данными
            clear_tables = {}
            # счётчик порядкового номера (количество) таблиц с очищенными данными
            number_clear_tables = 0
            # проходим по всем найденным таблицам (спискам списков)
            for i in data_tables:
                # проходим по всем спискам
                # счётчик (break_break) прерывания обхода списков, если нашли в ячейке "DIA", "Diameter",
                # "Nominal thickness" - для случая, если эти слова записаны в две объединённые строки
                break_break = 0
                for ii in data_tables[i]:
                    # проходим по всем элементам списка
                    for iii in ii:
                        # выбираем только таблицы с нужными данными в ячейках которых есть ключевые слова
                        # "DIA", "Diameter", "Nominal thickness"
                        if re.match(r'Nom', iii) or re.match(r'nom', iii) \
                                or re.match(r'DIA', iii) or re.match(r'Dia', iii):
                            clear_tables[number_clear_tables] = data_tables[i]
                            number_clear_tables += 1
                            break_break += 1
                            # если нашли ключевое слово, то прерываем дальнейший обход ячеек таблицы
                            # и переводим к следующей таблице
                            break
                    if break_break == 1:
                        break

            # очищенный (с начала таблицы) словарь с данными
            clear_table_top = {}
            # фильтруем полученные данные из таблиц, выбирая только данные, находящиеся ниже строки
            # "INSPECTION RESULTS", ключевого слова "Results"
            # переменная номера строки (INSPECTION RESULTS) в таблице после которой идут необходимые данные
            del_res_top = 0
            for i in list(clear_tables.keys()):
                for ii in range(len(clear_tables[i])):
                    for iii in range(len(clear_tables[i][ii])):
                        if re.findall(r'result|RESULT|Result', clear_tables[i][ii][0]):
                            del_res_top = ii + 1
                for iiii in range(del_res_top, len(clear_tables[i])):
                    clear_table_top[i] = clear_tables[i][del_res_top:]

            # очищенный (с конца таблицы) словарь с данными
            clear_table_bottom = {}
            # фильтруем отфильтрованную таблицу сверху (clear_table_top), выбирая только данные, находящиеся выше
            # строки "Examined by", ключевого слова "Exam"
            for i in list(clear_table_top.keys()):
                del_res_bottom = len(clear_table_top[i])
                for ii in range(len(clear_table_top[i])):
                    for iii in range(len(clear_table_top[i][ii])):
                        if re.findall(r'Examined by', clear_table_top[i][ii][iii]):
                            del_res_bottom = ii
                for iiii in range(del_res_bottom):
                    clear_table_bottom[i] = clear_table_top[i][:del_res_bottom]

            # очищаем полученные таблицы (clear_table_bottom) на данном этапе от пустых строк из-за наличия картинок
            for i in list(clear_table_bottom.keys()):
                # список номер пустых строк
                n_sp_str = []
                # активатор наличия пустых строк
                ch = 0
                for ii in range(len(clear_table_bottom[i])):
                    if len(set(clear_table_bottom[i][ii])) == 1 and '' in set(clear_table_bottom[i][ii]):
                        # обновляем список пустых строк
                        n_sp_str.append(ii)
                        # активируем счётчик наличия пустых строк
                        ch += 1
                if ch > 0:
                    p = 0
                    # если пустых строк больше, чем одна
                    if len(n_sp_str) > 1:
                        for d in n_sp_str:
                            # если первая пустая строка
                            if p == 0:
                                del clear_table_bottom[i][d]
                                p += 1
                            else:
                                # все последующие пустые строки
                                del clear_table_bottom[i][d - p]
                                p += 1
                    # если пустая строка только первая
                    elif len(n_sp_str) == 1:
                        del clear_table_bottom[i][n_sp_str[0]]

            # очищаем полученные таблицы (clear_table_bottom) на данном этапе от первой ошибочной таблицы из-за
            # наличия "Nominal Thickness", "Diameter" в самой первой таблице
            del_start_table = 0
            for i in list(clear_table_bottom.keys()):
                for ii in range(len(clear_table_bottom[i])):
                    for iii in range(len(clear_table_bottom[i][ii])):
                        if re.findall(r'Com', clear_table_bottom[i][ii][iii]) \
                                or re.findall(r'Proced', clear_table_bottom[i][ii][iii]):
                            del_start_table += 1
            if del_start_table > 0:
                del clear_table_bottom[0]

            # список пустых таблиц для сводки итоговых данных
            if clear_table_bottom == {}:
                # проверка, вносился ли уже репорт с таким номером в список
                if rep_number['report_number'] not in list_zero_table:
                    list_zero_table.append(rep_number['report_number'])
                    zero_table += 1

            # список репортов с нарушением структуры для итоговых данных
            for i in list(clear_table_bottom.keys()):
                try:
                    for ii in range(len(clear_table_bottom[i])):
                        # проверка, вносился ли уже репорт с таким номером в список
                        if rep_number['report_number'] not in list_distract_structure:
                            if len(clear_table_bottom[i][ii]) == 0:
                                list_distract_structure.append(rep_number['report_number'])
                                distract_structure += 1
                except KeyError:
                    mess = 'Ошибка ссылке на несуществующий ключ в словаре таблиц ' + rep_number['report_number']
                    message_mistake.append(mess)
                    dont_save_tables.append(rep_number['report_number'])
                    break

            # словарь названия столбцов для каждой таблицы
            name_column = {}
            # список номеров колонок с названиями линий
            number_column_line = []
            # список номеров колонок с названиями чертежей
            number_column_drawing = []
            # Выбираем названия столбцов для каждой таблицы
            for i in clear_table_bottom.keys():
                name_column[i] = clear_table_bottom[i][0]
                for ii in range(len(clear_table_bottom[i][0])):
                    if re.findall(r'Line|Tag', clear_table_bottom[i][0][ii]):
                        # номер колонки с названием линии для дальнейшего поиска
                        number_column_line.append(ii)
                    if re.findall(r'Draw|Isometr', clear_table_bottom[i][0][ii]):
                        # номер колонки с названием чертежа для дальнейшего поиска
                        number_column_drawing.append(ii)
            # удаляем из таблиц первые списки (названия столбцов)
            for i in clear_table_bottom.keys():
                del clear_table_bottom[i][0]

            # очищаем названия столбцов и приводим к допустимым названиям
            for i in name_column.keys():
                for ii in range(len(name_column[i])):
                    # проверяем наличие в названиях столбцов возможные недопустимые имена
                    if re.findall(r'Tag|TAG|Line|LINE', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Line')
                        # номер колонки с номером линии
                    elif re.findall(r'item|Item|description|Description', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Item_description')
                    elif re.findall(r'Section', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Section')
                    elif re.findall(r'Location', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Location')
                    elif re.findall(r'Remark', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Remark')
                    elif re.findall(r'Size|SIZE', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Size')
                    elif re.findall(r'Vertical', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Vertical')
                    elif re.findall(r'Horizontal', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Horizontal')
                    elif re.findall(r'Diametr|Dia|DIA|dia|DIА|Día', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Diameter')
                    elif re.findall(r'thicknes|Thicknes', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Nominal_thickness')
                    elif re.findall(r'Extrados', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Extrados')
                    elif re.findall(r'Intrados', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Intrados')
                    elif re.findall(r'South|south', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'South')
                    elif re.findall(r'North|north', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'North')
                    elif re.findall(r'West|west', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'West')
                    elif re.findall(r'East|east', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'East')
                    elif re.findall(r'Top|top', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Top')
                    elif re.findall(r'Bottom|bottom', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Bottom')
                    elif re.findall(r'Shell|shell', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Shell')
                    elif re.findall(r'Plate|plate', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Plate')
                    elif re.findall(r'Drawing|drawing|Isometr', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Drawing')
                    elif re.findall(r'Spot', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Spot')
                    elif re.findall(r'Cente', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Center')
                    elif re.findall(r'Row', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Row')
                    elif re.findall(r'Column', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Column')
                    elif re.findall(r'P&ID', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'P_ID')
                    elif re.findall(r'Right', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Right')
                    elif re.findall(r'Left', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Left')
                    elif re.findall(r'Date|date', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Date')
                    elif re.findall(r'Distance', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Distance')
                    elif re.findall(r'Result', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'Result')
                    elif re.findall(r'№|S/NO|S/N|s/n|s/no|NO|no', name_column[i][ii]):
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, 'S_N')
                    if re.findall(r'/', name_column[i][ii]):
                        b = re.sub(r'/', '_', name_column[i][ii])
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, b)
                    if re.findall(r' ', name_column[i][ii]):
                        b = re.sub(r' ', '_', name_column[i][ii])
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, b)
                    if re.findall(r'\n', name_column[i][ii]):
                        # меняем найденное значение
                        b = re.sub(r'\n', '', name_column[i][ii])
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, b)
                    if name_column[i][ii].isnumeric():
                        b = '_' + name_column[i][ii] + '_'
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, b)
                    if re.findall(r'\d+-\d+', name_column[i][ii]):
                        # меняем найденное значение
                        b = '_' + re.sub('-', '_', name_column[i][ii]) + '_'
                        # удаляем найденное значение
                        name_column[i].remove(name_column[i][ii])
                        # вставляем на удалённое место допустимое название столбца
                        name_column[i].insert(ii, b)

            # очищаем таблицу данных clear_table_bottom
            for i in list(clear_table_bottom.keys()):
                for ii in range(len(clear_table_bottom[i])):
                    for iii in range(len(clear_table_bottom[i][ii])):
                        if re.findall(r'"', clear_table_bottom[i][ii][iii]):
                            b = re.sub(r'"', '', clear_table_bottom[i][ii][iii])
                            clear_table_bottom[i][ii].remove(clear_table_bottom[i][ii][iii])
                            clear_table_bottom[i][ii].insert(iii, b)

            # создаём подключение к базе данных
            conn = sqlite3.connect('reports_db.db')
            cur = conn.cursor()
            for i in list(clear_table_bottom.keys()):
                # собираем название таблицы
                name_clear_table = '\'' + str(i) + '_' + rep_number['report_number'] + '\''
                try:
                    # создаем таблицу с именем name_clear_table и со столбцами name_column[i]
                    cur.execute('CREATE TABLE ' + name_clear_table + ' ({})'.format(','.join(name_column[i])))
                    conn.commit()
                except sqlite3.OperationalError:
                    mess = 'Ошибка в названии столбца (символ или дубль) ' + rep_number['report_number']
                    message_column.append(mess)
                    dont_save_tables.append(name_clear_table)
                    # сохраняем внесённые изменения, если не было ошибок в репорте
                    conn.commit()
                for ii in clear_table_bottom[i]:
                    try:
                        cur.execute('INSERT INTO ' + name_clear_table + ' VALUES (%s)' % ','.join('?' * len(ii)), ii)
                        conn.commit()
                    except sqlite3.OperationalError:
                        mess = 'Ошибка в названии столбца (символ или дубль) ' + rep_number['report_number']
                        message_column.append(mess)
                        dont_save_tables.append(name_clear_table)
                        # сохраняем внесённые изменения, если не было ошибок в репорте
                        conn.commit()
                if name_clear_table in dont_save_tables:
                    try:
                        cur.execute('DROP TABLE ' + name_clear_table)
                        conn.commit()
                    except sqlite3.OperationalError:
                        continue
                conn.commit()

            # закрываем соединение с базой данной
            conn.close()
    # сводка итоговых данных
    print('------------------------------------------------------------------------------------------------')
    print('Всего найдено файлов docx: ' + str(len(list_find_docx)))
    print('Из них репортов обработано: ' + str(len(list_find_docx) - len(set(list_cells)) -
                                               len(set(list_distract_structure)) - len(set(dont_save_tables))))
    print('Из них репортов с ошибками: ' + str(len(set(list_cells)) + len(set(list_distract_structure)) +
                                               len(set(dont_save_tables))))
    print('------------------------------------------------------------------------------------------------')
    print('Ошибки в репорте: ' + str(len(set(list_cells))))
    for i in range(len(message_mistake)):
        print('\t' + message_mistake[i])
    print('------------------------------------------------------------------------------------------------')
    print('Пустых таблиц: ' + str(len(set(list_zero_table))))
    print('В репорте:')
    for i in range(zero_table):
        print('\t' + list_zero_table[i])
    print('------------------------------------------------------------------------------------------------')
    print('Репортов с нарушением структуры таблицы: ' + str(len(set(list_distract_structure))))
    print('В репорте:')
    for i in range(distract_structure):
        print('\t' + list_distract_structure[i])
    print('------------------------------------------------------------------------------------------------')
    print('Количество репортов с ошибками в названиях столбцов: ' + str(len(set(dont_save_tables))))
    for i in range(len(set(dont_save_tables))):
        print('\t' + dont_save_tables[i])
    print('------------------------------------------------------------------------------------------------')


if __name__ == '__main__':
    add_table()
