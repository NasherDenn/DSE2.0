# -*- coding: utf-8 -*-

import os
import logging
import sys
import traceback
import YKR.utilities
# from YKR.utilities import number_report_wo_date
import datetime

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})

# настраиваем систему логирования
# дата (месяц, год) файла LogFile из системы
date_log_file = datetime.datetime.now().strftime("%m %Y")
# путь к папке где будет сохраняться LogFile
new_path_log_file = f'{os.path.abspath(os.getcwd())}\\Log File\\'
# если папка Log File не создана,
if not os.path.exists(new_path_log_file):
    # то создаём эту папку
    os.makedirs(new_path_log_file)
# путь сохранения LogFile
name_log_file = f'{new_path_log_file}{date_log_file} Log File.txt'
logging.basicConfig(level=logging.INFO,
                    handlers=[logging.FileHandler(filename=name_log_file, mode='a', encoding='utf-8')],
                    format='%(asctime)s [%(levelname)s] Пользователь: %(user)s - %(message)s', )
# дополняем базовый формат записи лог сообщения данными о пользователе
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})

logger_with_user.info('------------------------------------------------------------------------------------------------\n'
                      'Запуск программы')


def add_table():
    # для продакшн
    # def add_table(name_dir):

    # проверяем наличие всех БД (с 2019 по 2026 года) во всех вариациях в папке "DB"
    no_db_in_folder = YKR.utilities.db_in_folder()
    if no_db_in_folder:
        for i in no_db_in_folder:
            logger_with_user.error(f'В папке "DB" нет базы данных {i}')
        sys.exit('В папке "DB" нет базы данных ')

    # тест, надо будет поменять
    # путь к файлам для загрузки из диалогового окна выбора
    dir_files = 'C:/Users/Андрей/Documents/NDT/Тестовые данные/'

    # для продакшн
    name_dir = []
    for (dirpath, dirnames, filenames) in os.walk(dir_files):
        name_dir.extend(filenames)

    # тест
    # name_dir = ['04-YKR-OF-UTT-22-017 (Module-20, Loc-4,6,7).docx']

    # список путей и названий репортов для дальнейшей обработки
    list_name_reports_for_future_work = YKR.utilities.get_name_dir(dir_files, name_dir)
    # выбор только репортов в названиях которых есть "04-YKR"
    list_files_for_work = YKR.utilities.change_only_ykr_reports(list_name_reports_for_future_work)
    # начинаем перебирать репорты, прошедшие предварительную выборку
    for report in list_files_for_work:
        # переменная для перехода к следующему репорту, в случае выявленной ошибки
        break_break = True
        # получаем из первого верхнего колонтитула репорта неочищенные номер репорта, номер work order и дату
        dirty_rep_number = YKR.utilities.number_report_wo_date(report)
        # очищаем номер репорта, номер work order и дату от лишних (пробелы, новая строка) символов
        clear_rep_number = YKR.utilities.clear_data_rep_number(dirty_rep_number)
        # получаем название БД с локацией (ON, OF, OS), методом контроля (UTT, PAUT), годом контроля (18, 19, 20, 21, 22, 23, 24, 25, 26)
        name_reports_db, break_break = YKR.utilities.reports_db(clear_rep_number['report_number'], break_break)
        # если невозможно получить название БД из номера репорта, то переходим к следующему репорту с записью в Log File
        if not break_break:
            continue
        # извлекаем все таблицы из репорта в виде словарей
        dirty_data_report = YKR.utilities.get_dirty_data_report(report)
        # разделяем алгоритм, на "UTT" и "PAUT"
        if '-UTT-' in clear_rep_number['report_number'] or '-UT-' in clear_rep_number['report_number']:
            # первый перебор словарей (таблиц) в репорте
            # список номер словарей (таблиц) в которых есть ключевое слово "Nominal thickness"
            first_actual_table = []
            for number_dirty_table in dirty_data_report.keys():
                # выбираем только словари (таблицы) с данными
                # первый отбор по наличию в словаре (таблице) ключевого слова "Nominal thickness"
                first_actual_table.append(
                    YKR.utilities.first_clear_table_nominal_thickness(dirty_data_report[number_dirty_table], number_dirty_table,
                                                                      clear_rep_number['report_number']))
            # убираем None из первого отбора
            val = None
            first_actual_table = [i for i in first_actual_table if i != val]
            if not first_actual_table:
                report_number_for_logger = clear_rep_number['report_number']
                logger_with_user.warning(f'Не могу найти данные в репорте {report_number_for_logger}!')
            # словарь {"номер таблицы": "данные таблицы"}
            data_report_without_trash = {}
            if first_actual_table:
                # удаляем строку если она содержит "result", "details", "Notes
                for i in first_actual_table:
                    if type(i) == int:
                        data_report_without_trash[i] = YKR.utilities.delete_first_string(dirty_data_report[i])
            else:
                report_number_for_logger = clear_rep_number['report_number']
                logger_with_user.warning(
                    f'В репорте {report_number_for_logger} нет ключевого слова "Nominal thickness" или первая таблица с рабочей информацией не отделена '
                    f'от таблиц(ы) с данными!')
            # Проверяем таблицу, что бы в каждой строке было одинаковое количество ячеек.
            # Если нет, то в таблице есть сдвиги полей, т.е. таблица геометрически не ровная.
            data_table_equal_row = YKR.utilities.check_len_row(data_report_without_trash, clear_rep_number['report_number'])
            # убираем из дальнейшего перебора пустые данные
            if not data_table_equal_row:
                continue
            # определяем, какие номера таблиц являются "сеткой"
            mesh_table = YKR.utilities.which_table(data_table_equal_row)
            if mesh_table:
                # преобразуем таблицы с "сеткой" - переносим первые четыре строки в названия столбцов и их значения
                data_table_equal_row = YKR.utilities.converted_mesh(data_table_equal_row, mesh_table, clear_rep_number['report_number'])
            # data_table_equal_row - все таблицы на данном этапе, с том числе и преобразованная "сетка" в обычную
            # первый список (строка) - название столбцов
            # остальные списки (строки) - строки со значениями
            # приводим в порядок названия столбцов (первый список) и данные (остальные строки)
            # итоговый, очищенный, приведённый в порядок словарь pure_data_table = {"номер таблицы": [[названия столбцов], [[данные], [данные]]]}
            method = 'utt'
            pure_data_table = YKR.utilities.shit_in_shit_out(data_table_equal_row, method)
            # проверяем есть ли в столбце "Line" номер чертежа, если да, то разъединяем их и дополняем новым столбцом "Drawing"
            pure_data_table = YKR.utilities.check_drawing_in_line(pure_data_table)
            # проверяем и меняем повторяющиеся названия столбцов

            # ЗДЕСЬ ИТОГОВЫЕ ДАННЫЕ ДЛЯ ЗАПИСИ В БД
            pure_data_table = YKR.utilities.duplicate_name_column(pure_data_table)

            print(clear_rep_number['report_number'])
            print(pure_data_table)
            # return clear_rep_number['report_number'], pure_data_table

        if '-PAUT-' in clear_rep_number['report_number']:
            # первый перебор словарей (таблиц) в репорте
            # список номер словарей (таблиц) в которых есть ключевое слово "Nominal thickness"
            first_actual_table = []
            for number_dirty_table in dirty_data_report.keys():
                # выбираем только словари (таблицы) с данными
                # первый отбор по наличию в словаре (таблице) ключевого слова "Nominal thickness"
                first_actual_table.append(
                    YKR.utilities.first_clear_table_nominal_thickness(dirty_data_report[number_dirty_table], number_dirty_table,
                                                                      clear_rep_number['report_number']))
            # убираем None из первого отбора
            val = None
            first_actual_table = [i for i in first_actual_table if i != val]
            if not first_actual_table:
                report_number_for_logger = clear_rep_number['report_number']
                logger_with_user.warning(f'Не могу найти данные в репорте {report_number_for_logger}!')
            # словарь {"номер таблицы": "данные таблицы"}
            data_report_without_trash = {}
            if first_actual_table:
                # удаляем строку если она содержит "result", "details", "Notes
                for i in first_actual_table:
                    if type(i) == int:
                        data_report_without_trash[i] = YKR.utilities.delete_first_string(dirty_data_report[i])
            else:
                report_number_for_logger = clear_rep_number['report_number']
                logger_with_user.warning(
                    f'В репорте {report_number_for_logger} нет ключевого слова "Nominal thickness" или первая таблица с рабочей информацией не отделена '
                    f'от таблиц(ы) с данными!')
            # Проверяем таблицу, что бы в каждой строке было одинаковое количество ячеек.
            # Если нет, то в таблице есть сдвиги полей, т.е. таблица геометрически не ровная.
            data_table_equal_row = YKR.utilities.check_len_row(data_report_without_trash, clear_rep_number['report_number'])
            # убираем из дальнейшего перебора пустые данные
            if not data_table_equal_row:
                continue
            method = 'paut'
            # итоговый, очищенный, приведённый в порядок словарь pure_data_table = {"номер таблицы": [[названия столбцов], [[данные], [данные]]]}
            pure_data_table = YKR.utilities.shit_in_shit_out(data_table_equal_row, method)
            # проверяем есть ли в столбце "Line" номер чертежа, если да, то разъединяем их и дополняем новым столбцом "Drawing"

            # ЗДЕСЬ ИТОГОВЫЕ ДАННЫЕ ДЛЯ ЗАПИСИ В БД
            pure_data_table = YKR.utilities.check_drawing_in_line(pure_data_table)

            print(clear_rep_number['report_number'])
            print(pure_data_table)
            # return clear_rep_number['report_number'], pure_data_table


def main():
    add_table()


if __name__ == '__main__':
    main()
