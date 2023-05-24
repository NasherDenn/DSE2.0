# -*- coding: utf-8 -*-

# import os
# import sys
# import logging
import itertools


# проверка выбраны ли фильтры и введены ли данные для поиска
def not_check_data_and_filter(data_filter_for_search: dict, data_for_search: dict) -> bool:
    check_filter_location = False
    check_filter_method = False
    check_filter_year = False
    data = False
    for location in data_filter_for_search['location']:
        if data_filter_for_search['location'][location]:
            check_filter_location = True
    for method in data_filter_for_search['method']:
        if data_filter_for_search['method'][method]:
            check_filter_method = True
    for year in data_filter_for_search['year']:
        if data_filter_for_search['year'][year]:
            check_filter_year = True
    for key in data_for_search.keys():
        if data_for_search[key][0] != '':
            data = True
    if check_filter_location and check_filter_method and check_filter_year and data:
        return False
    else:
        return True


# названия БД по выбранным фильтрам в которых надо искать данные
def define_db_for_search(data_filter_for_search: dict) -> list:
    db_for_search = list()
    location = list()
    method = list()
    year = list()
    # выбираем выбранные фильтры
    for check_filter in data_filter_for_search['location']:
        if data_filter_for_search['location'][check_filter]:
            location.append(check_filter)
    for check_filter in data_filter_for_search['method']:
        if data_filter_for_search['method'][check_filter]:
            method.append(check_filter)
    for check_filter in data_filter_for_search['year']:
        if data_filter_for_search['year'][check_filter]:
            year.append(check_filter[2:])
    all_combinations = list(itertools.product(location, method, year))
    for combination in all_combinations:
        db_for_search.append(f'reports_db_{combination[0]}_{combination[1]}_{combination[2]}.sqlite')
    return db_for_search
