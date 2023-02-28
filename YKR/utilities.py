from docx import Document
import re
import logging
import traceback
import os

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})


# получение пути и названий репортов для дальнейшей работы
def get_name_dir(name_dir_files: list) -> list:
    # переменная-список для дальнейшего преобразования списка списков в список строк выбранных для загрузки файлов docx
    name_dir_docx = []
    for i in name_dir_files:
        name_dir_docx.append(f'C:/Users/asus/Documents/NDT YKR/Тестовые данные/{i}')
    return name_dir_docx


# выбор только файлов (репортов) в названиях которых есть "04-YKR"
def change_only_ykr_reports(name_dir_docx: list) -> list:
    list_reports_for_work = []
    for i in name_dir_docx:
        if '04-YKR' in i:
            list_reports_for_work.append(i)
    return list_reports_for_work


# получаем номер репорта, номер work order и дату
def number_report_wo_date(path_to_report: str) -> dict:
    doc = Document(path_to_report)
    # получаем неочищенные данные из первого верхнего колонтитула
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
    return get_number_report_wo_date(data_header)


# получение фактического номера репорта, номера work order и даты
def get_number_report_wo_date(data_header: dict) -> dict:
    # словарь номера репорта, даты репорта, номера Work Order
    rep_number = {}
    # перебираем полученные данные из первого верхнего колонтитула
    # проходим по нему как по таблице
    for index_row, i in enumerate(data_header[0]):
        for index_column, ii in enumerate(i):
            # если слова "Report No:" и фактический номер репорта в одной ячейке
            if 'Report' in ii and any(map(str.isdigit, data_header[0][index_row][index_column])):
                # индекс начала номера репорта в той же ячейке
                index_start_number_report = re.search("\d", ii).start()
                rep_number['report_number'] = data_header[0][index_row][index_column][index_start_number_report:]
            # иначе если слова "Report No:" и фактический номер репорта в разных ячейках
            elif 'Report' in ii and any(map(str.isdigit, data_header[0][index_row][index_column + 1])):
                # индекс начала номера репорта в соседней ячейке
                index_start_number_report = re.search("\d", data_header[0][index_row][index_column + 1]).start()
                rep_number['report_number'] = data_header[0][index_row][index_column + 1][index_start_number_report:]
            # если слово "Date" и фактическая дата репорта в одной ячейке
            if 'Date' in ii and any(map(str.isdigit, data_header[0][index_row][index_column])):
                # индекс начала даты репорта в той же ячейке
                index_start_number_report = re.search("\d", ii).start()
                rep_number['report_date'] = data_header[0][index_row][index_column][index_start_number_report:]
            # иначе если слово "Date" и фактическая дата репорта в разных ячейках
            elif 'Date' in ii and any(map(str.isdigit, data_header[0][index_row][index_column + 1])):
                # индекс начала даты репорта в соседней ячейке
                index_start_number_report = re.search("\d", data_header[0][index_row][index_column + 1]).start()
                rep_number['report_date'] = data_header[0][index_row][index_column + 1][index_start_number_report:]
            # если слово "order" и номер work order в одной ячейке
            if 'order' in ii and any(map(str.isdigit, data_header[0][index_row][index_column])):
                # индекс начала номера work order в той же ячейке
                index_start_number_report = re.search("\d", ii).start()
                rep_number['work_order'] = data_header[0][index_row][index_column][index_start_number_report:]
            # иначе если слова "order" и номер work order в разных ячейках
            elif 'order' in ii and any(map(str.isdigit, data_header[0][index_row][index_column + 1])):
                # индекс начала номера work order в соседней ячейке
                index_start_number_report = re.search("\d", data_header[0][index_row][index_column + 1]).start()
                rep_number['work_order'] = data_header[0][index_row][index_column + 1][index_start_number_report:]
            # иначе если слова "order" и номер work order в разных ячейках и нет цифр, значит номер work order - NCOC Request
            elif 'order' in ii and not any(map(str.isdigit, data_header[0][index_row][index_column + 1])):
                rep_number['work_order'] = data_header[0][index_row][index_column + 1]
    # возвращаем не очищенные значения номера репорта, номера work order и даты
    return rep_number


# очистка номера репорта, даты репорта, номера Work Order от лишних, повторяющихся символов
def clear_data_rep_number(data: dict) -> dict:
    # удаление любых пробельных символов в номере репорта
    data['report_number'] = re.sub('\s+', '', data['report_number'])
    # замена повторяющегося символа "-" на единичный в номере репорта
    data['report_number'] = re.sub('-+', '-', data['report_number'])
    # замена любых пробельных символов в дате репорта на "."
    data['report_date'] = re.sub('\s+', '.', data['report_date'])
    # замена повторяющегося символа "." на единичный в дате репорта
    data['report_date'] = re.sub('\.+', '.', data['report_date'])
    # замена повторяющегося символа "-" на единичный в дате репорта
    data['report_date'] = re.sub('-+', '-', data['report_date'])
    # удаление любых пробельных символов в work order
    data['work_order'] = re.sub('\s+', '', data['work_order'])
    # если в номере репорта была Revision
    if 'Rev' in data['report_number'] or 'rev' in data['report_number'] or 'REV' in data['report_number']:
        # то добавляем "Rev." через знак "_"
        index_rev = data['report_number'].find('ev')
        data['report_number'] = '_'.join([data['report_number'][:index_rev - 1], data['report_number'][index_rev - 1:]])
    print(data)
    return data


# вытягиваем из номера репорта локацию (ON, OF, OS), метод контроля (UTT, PAUT), год контроля (18, 19, 20, 21, 22, 23, 24, 25, 26)
# и формирование имени БД для дальнейшей записи
def reports_db(name_report: str, break_break: bool) -> tuple:
    location = ['-ON-', '-on-', '-OF-', '-of-', '-OFF-', '-off-', '-OS-', '-os-' ]
    method = ['-UT-', '-ut-', '-UTT-', '-utt-', '-PAUT-', '-paut-']
    years = ['-18-', '-19-', '-20-', '-21-', '-22-', '-23-', '-24-', '-25-', '-26-']
    name_for_reports_db = ''
    if break_break:
        for i in location:
            if i in name_report:
                name_for_reports_db = f'reports_db_{i[1:-1]}_'
                break
        if name_for_reports_db == '':
            name_for_reports_db = 'reports_db_ON_'
        # активатор, если не нашли метод контроля в номере репорта
        find = False
        for i in method:
            if i in name_report:
                name_for_reports_db = f'{name_for_reports_db}{i[1:-1].upper()}_'
                if '_UT_' in name_for_reports_db:
                    name_for_reports_db = name_for_reports_db.replace('_UT_', '_UTT_')
                if '_OF_' in name_for_reports_db:
                    name_for_reports_db = name_for_reports_db.replace('_OF_', '_OFF_')
                find = True
        # если не нашли метод контроля, то переходим к следующему репорту
        if not find:
            logger_with_user.error(f'Не могу определить метод контроля! Проверь корректность записи номера репорта {name_report}!1')
            break_break = False
        if break_break:
            find = False
            for i in years:
                if i in name_report:
                    name_for_reports_db = f'{name_for_reports_db}{i[1:-1]}.sqlite'
                    find = True
            if not find:
                logger_with_user.error(f'Не могу определить год контроля! Проверь корректность записи номера репорта {name_report}!2')
                break_break = False
    return name_for_reports_db, break_break


def get_dirty_data_report(path_to_report: str) -> dict:
    doc = Document(path_to_report)
    # переменная со всеми таблицами в репорте
    all_tables = doc.tables
    # создаем пустой словарь под неочищенные данные таблиц
    dirty_data_tables = {i: None for i in range(0, len(all_tables))}
    for i, table in enumerate(all_tables):
        # создаем список строк для таблицы `i` (пока пустые)
        dirty_data_tables[i] = [[] for _ in range(0, len(table.rows))]
        # проходимся по строкам таблицы `i`
        for j, row in enumerate(table.rows):
            for cell in row.cells:
                dirty_data_tables[i][j].append(cell.text)
    # print(dirty_data_tables)
    return dirty_data_tables


# получаем только словари (таблицы) в которых есть ключевое слово "Nominal thickness"
def first_clear_table_nominal_thickness(first_dirty_table: dict, number_dirty_table: int) -> int:
    # перебираем строки в словаре (таблице)
    for row in first_dirty_table:
        # перебираем колонки в строке
        for column in row:
            if 'Nom' in column or 'nom' in column or 'NOM' in column:
                return number_dirty_table


# получаем словари (таблицы) в которых есть слово "Project"
def second_clear_table_mistake_first_table(dirty_data_report: dict, first_actual_table: int) -> int:
    for row in dirty_data_report[first_actual_table]:
        for column in row:
            if 'Proj' in column:
                return first_actual_table
