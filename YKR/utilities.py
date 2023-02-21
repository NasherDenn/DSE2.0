from docx import Document
import re


# получение пути и названий репортов для дальнейшей работы
def get_name_dir(name_dir_files: list) -> list:
    # переменная-список для дальнейшего преобразования списка списков в список строк выбранных для загрузки файлов docx
    name_dir_docx = []
    for i in name_dir_files:
        name_dir_docx.append(r'C:/Users/asus/Documents/NDT YKR/Тестовые данные/' + i)
    return name_dir_docx


# выбор только файлов (репортов) в названиях которых есть "04-YKR"
def change_only_ykr_reports(name_dir_docx: list) -> list:
    list_reports_for_work = []
    for i in name_dir_docx:
        if '04-YKR' in i:
            list_reports_for_work.append(i)
    return list_reports_for_work


# получаем номер репорта, номер work order и дату
def number_report_wo_date(path_to_report: str) -> tuple:
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
    return get_number_report_wo_date(data_header), doc


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
                # индекс начала номера репорта в той же ячейке
                index_start_number_report = re.search("\d", ii).start()
                rep_number['report_date'] = data_header[0][index_row][index_column][index_start_number_report:]
            # иначе если слова "Report No:" и фактический номер репорта в разных ячейках
            elif 'Date' in ii and any(map(str.isdigit, data_header[0][index_row][index_column + 1])):
                # индекс начала номера репорта в соседней ячейке
                index_start_number_report = re.search("\d", data_header[0][index_row][index_column + 1]).start()
                rep_number['report_date'] = data_header[0][index_row][index_column + 1][index_start_number_report:]
            # если слово "order" и номер work order в одной ячейке
            if 'order' in ii and any(map(str.isdigit, data_header[0][index_row][index_column])):
                # индекс начала номера репорта в той же ячейке
                index_start_number_report = re.search("\d", ii).start()
                rep_number['work_order'] = data_header[0][index_row][index_column][index_start_number_report:]
            # иначе если слова "order" и номер work order в разных ячейках
            elif 'order' in ii and any(map(str.isdigit, data_header[0][index_row][index_column + 1])):
                # индекс начала номера репорта в соседней ячейке
                index_start_number_report = re.search("\d", data_header[0][index_row][index_column + 1]).start()
                rep_number['work_order'] = data_header[0][index_row][index_column + 1][index_start_number_report:]
    # возвращаем не очищенные значения номера репорта, номера work order и даты
    return rep_number


# очистка номера репорта, даты репорта, номера Work Order от лишних (пробелы, новая строка) символов
def clear_data_rep_number(data: dict) -> dict:
    # от любых пробельных символов
    data[0]['report_number'] = re.sub('\s+', '', data[0]['report_number'])
    data[0]['report_date'] = re.sub('\s+', '.', data[0]['report_date'])
    data[0]['work_order'] = re.sub('\s+', '', data[0]['work_order'])
    # если в номере репорта была Revision
    if 'Rev' in data[0]['report_number'] or 'rev' in data[0]['report_number'] or 'REV' in data[0]['report_number']:
        # номер начала слова "Rev." в номере репорта
        index_rev = data[0]['report_number'].find('ev')
        data[0]['report_number'] = '_'.join([data[0]['report_number'][:index_rev - 1], data[0]['report_number'][index_rev - 1:]])
    print(data[0])
    return data[0]


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
def first_clear_table_nominal_thickness(first_dirty_table, number_dirty_table):
    # перебираем строки в словаре (таблице)
    for row in first_dirty_table:
        # перебираем колонки в строке
        for column in row:
            if 'Nom' in column or 'nom' in column or 'NOM' in column:
                return number_dirty_table


# получаем словари (таблицы) в которых есть слово "Project"
def second_clear_table_mistake_first_table(dirty_data_report, first_actual_table):
    for row in dirty_data_report[first_actual_table]:
        for column in row:
            if 'Proj' in column:
                return first_actual_table
