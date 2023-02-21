import os
import logging
import traceback
import utilities
from utilities import number_report_wo_date

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})


def add_table():
    # def add_table(name_dir):

    name_dir = []
    for (dirpath, dirnames, filenames) in os.walk(r'C:/Users/asus/Documents/NDT YKR/Тестовые данные/'):
        name_dir.extend(filenames)

    # список путей и названий репортов для дальнейшей обработки
    list_name_reports_for_future_work = utilities.get_name_dir(name_dir)
    # выбор только репортов в названиях которых есть "04-YKR"
    list_files_for_work = utilities.change_only_ykr_reports(list_name_reports_for_future_work)
    # начинаем перебирать репорты, прошедшие предварительную выборку
    for report in list_files_for_work:
        # получаем из первого верхнего колонтитула репорта не очищенные номер репорта, номер work order и дату
        dirty_rep_number = utilities.number_report_wo_date(report)
        # очищаем номер репорта, номер work order и дату от лишних (пробелы, новая строка) символов
        clear_rep_number = utilities.clear_data_rep_number(dirty_rep_number)
        # извлекаем все таблицы из репорта в виде словарей
        dirty_data_report = utilities.get_dirty_data_report(report)
        # первый перебор словарей (таблиц) в репорте
        # список номер словарей (таблиц) в которых есть ключевое слово "Nominal thickness"
        first_actual_table = []
        for number_dirty_table in dirty_data_report.keys():
            # выбираем только словари (таблицы) с данными
            # первый отбор по наличию в словаре (таблице) ключевого слова "Nominal thickness"
            first_actual_table.append(
                utilities.first_clear_table_nominal_thickness(dirty_data_report[number_dirty_table], number_dirty_table))

        # убираем None из первого отбора
        val = None
        first_actual_table = [i for i in first_actual_table if i != val]

        # вторым отбором получаем номера ошибочных (первых) таблиц из-за наличия в них "Nominal Thickness"
        print(first_actual_table)
        # список ошибочных номеров таблиц
        table_with_mistake_nominal_thickness = []
        if first_actual_table:
            for number_first_clear_table in first_actual_table:
                table_with_mistake_nominal_thickness.append(utilities.second_clear_table_mistake_first_table(dirty_data_report, number_first_clear_table))
        print(table_with_mistake_nominal_thickness)
        for table in table_with_mistake_nominal_thickness:
            if table in first_actual_table:
                first_actual_table.remove(table)
        print(first_actual_table)


def main():
    add_table()


if __name__ == '__main__':
    main()
