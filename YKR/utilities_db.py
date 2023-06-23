import logging
import os
import sqlite3
import traceback
# import sys
# import re


# from PyQt5.QtWidgets import *

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})


# записываем очищенный репорт в базу данных
# clear_report - очищенные, переименованные таблицы
# clear_report - номер репорта
# name_db - имя БД для записи
def write_report_in_db(clear_report: dict, number_report: dict, name_db: str, first_actual_table: list, unit: list):
    # меняем все "-" на "_" что бы записать в БД
    true_number_report = number_report['report_number'].replace('-', '_')
    true_number_report = true_number_report.replace('.', '_')
    # подключаемся к БД
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
    # включаем поддержку внешних ключей
    conn.execute('''PRAGMA foreign_keys = ON''')
    conn.commit()
    cur = conn.cursor()
    # включаем поддержку внешних ключей
    # cur.execute('''PRAGMA foreign_keys = ON''')

    # создаём таблицу master со столбцами из clear_rep_number, количества таблиц, списка таблиц ЕСЛИ ещё не существует
    if not cur.execute('''SELECT * FROM sqlite_master WHERE type="table" AND name="master"''').fetchall():
        cur.execute('''CREATE TABLE IF NOT EXISTS master (unit, report_number PRIMARY KEY, report_date, work_order, one_of, list_table_report)''')
        conn.commit()
        # создаём индекс
        cur.execute('''CREATE INDEX id ON master (unit)''')
        conn.commit()
    for number_table in clear_report.keys():
        # can_write_rep_number_in_master = False
        # собираем имя таблицы для записи
        name_table_for_write = f'_{number_table}_{true_number_report}'

        write_rep_number_in_master(number_report, first_actual_table, name_table_for_write, name_db, unit)

        if not cur.execute('''SELECT * FROM sqlite_master WHERE  tbl_name="{}"'''.format(name_table_for_write)).fetchone():
            try:
                # добавляем в названия столбцов на первое место number_report
                # cur.execute('''CREATE TABLE IF NOT EXISTS {} ({})'''.format(name_table_for_write, ','.join(clear_report[number_table][0])))

                rep = f'number_report, {",".join(clear_report[number_table][0])}, FOREIGN KEY(number_report) REFERENCES master(report_number) ON UPDATE CASCADE'
                # rep = f'number_report, {",".join(clear_report[number_table][0])}, FOREIGN KEY(number_report) REFERENCES master(report_number)'
                print(name_table_for_write)
                print(rep)
                cur.execute('''CREATE TABLE IF NOT EXISTS {} ({})'''.format(name_table_for_write, rep))
                # FOREIGN KEY(k) REFERENCES sss(a)
                # print('ok')


                conn.commit()
            except sqlite3.OperationalError:
                logger_with_user.error(f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
                                       f'{traceback.format_exc()}')
                # no such table: main.master

                continue
            for values in clear_report[number_table][1]:
                try:
                    # print(values)
                    values.insert(0, number_report["report_number"])
                    # print(values)

                    # print(number_report["report_number"])

                    cur.execute('INSERT INTO ' + name_table_for_write + ' VALUES (%s)' % ','.join('?' * len(values)), values)
                    conn.commit()
                    # can_write_rep_number_in_master = True

                except sqlite3.OperationalError:
                    logger_with_user.error(f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
                                           f'{traceback.format_exc()}')
                    continue
        # если таблица удачно записана в БД, то записываем номер репорта, wo, дату в таблицу master
        # number_report - словарь номера репорта, даты, wo
        # first_actual_table - словарь номеров таблиц в которых есть необходимые данные
        # name_table_for_write - имя таблицы
        # name_db - имя БД для записи

        # if can_write_rep_number_in_master:
        #     write_rep_number_in_master(number_report, first_actual_table, name_table_for_write, name_db, unit)
    cur.close()


# запись в таблицу master unit, номера репорта, wo, даты, количества таблиц в репорте
def write_rep_number_in_master(number_report: dict, count_table: list, name_table: str, name_db: str, unit: list):
    # форматируем номер таблицы для лучшей визуализации (меняем "_" на "-")
    name_table = name_table.replace("_", "-")[1:]
    # подключаемся к БД
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
    cur = conn.cursor()


    # # создаём таблицу master со столбцами из clear_rep_number, количества таблиц, списка таблиц ЕСЛИ ещё не существует
    # if not cur.execute('''SELECT * FROM sqlite_master WHERE type="table" AND name="master"''').fetchall():
    #     cur.execute('''CREATE TABLE IF NOT EXISTS master (unit, report_number, report_date, work_order, one_of, list_table_report)''')
    #     conn.commit()
    #     # создаём индекс
    #     cur.execute('''CREATE INDEX id ON master (unit)''')
    #     conn.commit()


    # если в master нет такого номера репорта, то записываем его первым со значение one_of (1/...) и номером таблицы
    if not cur.execute('''SELECT * FROM master WHERE report_number="{}"'''.format(number_report['report_number'])).fetchone():
        cur.execute('INSERT INTO master VALUES (?, ?, ?, ?, ?, ?)',
                    (unit,
                     number_report['report_number'],
                     number_report['report_date'],
                     number_report['work_order'],
                     f'1/{len(count_table)}',
                     name_table))
        conn.commit()
    # иначе проверяем номер unit
    else:
        # если номер unit такой же, то обновляем строчку с записью в master
        if unit == cur.execute('''SELECT unit FROM master WHERE report_number = "{}"'''.format(number_report['report_number'])).fetchall()[0][0]:
            # пересчитываем и переписываем one_of и дописываем list_table_report номером новой таблицы
            old_one_of = cur.execute('''SELECT one_of FROM master WHERE report_number = "{}"'''.format(number_report['report_number'])).fetchall()[0][0]
            index_slash_old_one_of = old_one_of.find('/')
            old_one_of_left_slash = int(old_one_of[:index_slash_old_one_of])
            new_one_of_left_slash = str(old_one_of_left_slash + 1)
            update_one_of = f'{new_one_of_left_slash}{old_one_of[index_slash_old_one_of:]}'
            old_list_table_report = cur.execute('''SELECT list_table_report FROM master WHERE report_number = "{}"'''
                                                .format(number_report['report_number'])).fetchall()[0][0]
            update_list_table_report = f'{old_list_table_report}\n{name_table}'
            # print(update_one_of)
            # print(update_list_table_report)
            # print(number_report['report_number'])
            # print(unit)
            cur.execute('''UPDATE  master set one_of='{}', list_table_report='{}' WHERE report_number="{}" AND unit="{}"'''
                        .format(update_one_of,
                                update_list_table_report,
                                number_report['report_number'],
                                unit))
            conn.commit()
        # иначе записываем новую строчку
        else:
            try:
                cur.execute('INSERT INTO master VALUES (?, ?, ?, ?, ?, ?)',
                            (unit,
                             number_report['report_number'],
                             number_report['report_date'],
                             number_report['work_order'],
                             f'1/{len(count_table)}',
                             name_table))
                conn.commit()
            except sqlite3.IntegrityError:
                logger_with_user.error(f'Проверь данные для записи в репорте {number_report}.\n'
                                       f'{traceback.format_exc()}')
    cur.close()


# ищем данные в БД
# db_for_search - список БД, в которых надо искать
# values_for_search - словарь с введёнными значениями для поиска в соответствующее поле (номер линии, номер чертежа, номер локации, номер отчёта)
def look_up_data(db_for_search: list, values_for_search: dict, search_path: int):
    for name_db in db_for_search:
        # подключаемся к БД
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
        cur = conn.cursor()
        # список всех таблиц в БД
        list_table_for_search = cur.execute('''SELECT name FROM sqlite_master WHERE type='table';''').fetchall()
        for table in list_table_for_search:
            # если заполнена одна строка в поле для поиска
            if table[0] != 'master' and search_path == 1:
                if not values_for_search['number_report']:
                    for i in values_for_search.keys():
                        if values_for_search[i]:
                            try:
                                if cur.execute('''SELECT * FROM {} WHERE {} LIKE "%{}"'''.format(table[0], i, values_for_search[i])).fetchall():
                                    return cur.execute('''SELECT * FROM {} WHERE {} LIKE "%{}"'''.format(table[0], i, values_for_search[i])).fetchall()
                            except:
                                continue
