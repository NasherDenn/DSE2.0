# -*- coding: utf-8 -*-

import os
import sys
import logging
import datetime

# from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
# from PyQt5.QtSql import QSqlDatabase
from PyQt5.QtSql import QSqlQueryModel, QSqlTableModel
from PyQt5.QtWidgets import *
from YKR.utilities_interface import *
from YKR.utilities_db import *

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

# создаём приложение
app = QApplication(sys.argv)
# создаём окно приложения
window = QWidget()
# название приложения
window.setWindowTitle('Data Search Engine')
# задаём стиль приложения Fusion
app.setStyle('Fusion')
# размер окна приложения
window.setFixedSize(1722, 965)

# устанавливаем favicon в окне приложения
icon = QIcon()
icon.addFile(u"icon.ico", QSize(), QIcon.Normal, QIcon.Off)
icon.addFile(u"icon.ico", QSize(), QIcon.Active, QIcon.On)

app.setWindowIcon(icon)

# задаём параметры стиля и оформления окна ввода
font = QFont()
font.setFamily(u"Arial")
font.setPointSize(14)
font.setItalic(False)

# создаём однострочное поле для ввода номера линии
line_search_line = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_line.setGeometry(QRect(181, 20, 561, 31))
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_line.setObjectName(u"line_search_line")
# дополнительные параметры
line_search_line.setFont(font)
line_search_line.setMouseTracking(False)
line_search_line.setFocusPolicy(Qt.ClickFocus)
line_search_line.setContextMenuPolicy(Qt.NoContextMenu)
line_search_line.setAcceptDrops(True)
line_search_line.setStyleSheet(u"")
line_search_line.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_line.setEchoMode(QLineEdit.Normal)
line_search_line.setCursorPosition(0)
line_search_line.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_line.setClearButtonEnabled(True)
line_search_line.setText('28278087')
line_search_line.setFocus()

# создаём однострочное поле для ввода номера чертежа
line_search_drawing = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_drawing.setGeometry(QRect(181, 60, 561, 31))
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_drawing.setObjectName(u"line_search_drawing")
# дополнительные параметры
line_search_drawing.setFont(font)
line_search_drawing.setMouseTracking(False)
line_search_drawing.setFocusPolicy(Qt.ClickFocus)
line_search_drawing.setContextMenuPolicy(Qt.NoContextMenu)
line_search_drawing.setAcceptDrops(True)
line_search_drawing.setStyleSheet(u"")
line_search_drawing.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_drawing.setEchoMode(QLineEdit.Normal)
line_search_drawing.setCursorPosition(0)
line_search_drawing.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_drawing.setClearButtonEnabled(True)
line_search_drawing.setText('')

# создаём однострочное поле для ввода номера юнита
line_search_unit = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_unit.setGeometry(QRect(181, 100, 170, 31))
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_unit.setObjectName(u"line_search_unit")
# дополнительные параметры
line_search_unit.setFont(font)
line_search_unit.setMouseTracking(False)
line_search_unit.setFocusPolicy(Qt.ClickFocus)
line_search_unit.setContextMenuPolicy(Qt.NoContextMenu)
line_search_unit.setAcceptDrops(True)
line_search_unit.setStyleSheet(u"")
line_search_unit.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_unit.setEchoMode(QLineEdit.Normal)
line_search_unit.setCursorPosition(0)
line_search_unit.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_unit.setClearButtonEnabled(True)
line_search_unit.setText('')

# создаём однострочное поле для ввода номера локации
line_search_item_description = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_item_description.setGeometry(QRect(521, 100, 220, 31))
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_item_description.setObjectName(u"line_search_item_description")
# дополнительные параметры
line_search_item_description.setFont(font)
line_search_item_description.setMouseTracking(False)
line_search_item_description.setFocusPolicy(Qt.ClickFocus)
line_search_item_description.setContextMenuPolicy(Qt.NoContextMenu)
line_search_item_description.setAcceptDrops(True)
line_search_item_description.setStyleSheet(u"")
line_search_item_description.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_item_description.setEchoMode(QLineEdit.Normal)
line_search_item_description.setCursorPosition(0)
line_search_item_description.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_item_description.setClearButtonEnabled(True)
line_search_item_description.setText('')

# создаём однострочное поле для ввода номера репорта
line_search_number_report = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_number_report.setGeometry(QRect(181, 140, 561, 31))
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_number_report.setObjectName(u"line_search_number_report")
# дополнительные параметры
line_search_number_report.setFont(font)
line_search_number_report.setMouseTracking(False)
line_search_number_report.setFocusPolicy(Qt.ClickFocus)
line_search_number_report.setContextMenuPolicy(Qt.NoContextMenu)
line_search_number_report.setAcceptDrops(True)
line_search_number_report.setStyleSheet(u"")
line_search_number_report.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_number_report.setEchoMode(QLineEdit.Normal)
line_search_number_report.setCursorPosition(0)
line_search_number_report.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_number_report.setClearButtonEnabled(True)
line_search_number_report.setText('')

# создаём кнопку "Поиск"
button_search = QPushButton('Поиск', window)
# устанавливаем положение и размер кнопки для поиска в родительском окне (window)
button_search.setGeometry(760, 20, 161, 41)
# присваиваем уникальное объектное имя кнопке "Поиск"
button_search.setObjectName(u"pushButton_search")
# задаём параметры стиля и оформления кнопки "Поиск"
font_button_search = QFont()
font_button_search.setFamily(u"Arial")
font_button_search.setPointSize(14)
button_search.setFont(font_button_search)
# дополнительные параметры
button_search.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку печати
button_print = QPushButton('Печать', window)
# устанавливаем положение и размер кнопки печати в родительском окне (window)
button_print.setGeometry(QRect(760, 70, 161, 41))
# присваиваем уникальное объектное имя кнопке "Печать"
button_print.setObjectName(u"pushButton_print")
# задаём параметры стиля и оформления кнопки печати
font_button_print = QFont()
font_button_print.setFamily(u"Arial")
font_button_print.setPointSize(14)
button_print.setFont(font_button_print)
# дополнительные параметры
button_print.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку "Добавить"
button_add = QPushButton('Добавить', window)
# устанавливаем положение и размер кнопки "Добавить" в родительском окне (window)
button_add.setGeometry(QRect(760, 120, 161, 41))
# присваиваем уникальное объектное имя кнопке "Добавить"
button_add.setObjectName(u"pushButton_add")
# задаём параметры стиля и оформления кнопки "Добавить"
font_button_add = QFont()
font_button_add.setFamily(u"Arial")
font_button_add.setPointSize(14)
button_add.setFont(font_button_add)
# дополнительные параметры
button_add.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Добавить" до авторизации
button_add.setDisabled(True)

# создаём кнопку "Закрыть" из программы
button_exit = QPushButton('Закрыть', window)
# устанавливаем положение и размер кнопки "Закрыть" для выхода из программы в родительском окне (window)
button_exit.setGeometry(QRect(1540, 904, 161, 41))
# присваиваем уникальное объектное имя кнопке "Закрыть"
button_exit.setObjectName(u"pushButton_exit")
# задаём параметры стиля и оформления кнопки "Закрыть"
font_button_exit = QFont()
font_button_exit.setFamily(u"Arial")
font_button_exit.setPointSize(14)
button_exit.setFont(font_button_exit)
# дополнительные параметры
button_exit.setFocusPolicy(Qt.ClickFocus)
# Закрытие программы при нажатии на кнопку "Закрыть"
button_exit.clicked.connect(qApp.exit)

# создаём однострочное поле для ввода логина
line_login = QLineEdit(window)
# присваиваем уникальное объектное имя полю для ввода логина
line_login.setObjectName(u"line_login")
# устанавливаем положение и размер поля для ввода логина в родительском окне (window)
line_login.setGeometry(QRect(1470, 20, 111, 31))
# задаём параметры стиля и оформления поля для ввода логина
font_line_login = QFont()
font_line_login.setFamily(u"Arial")
font_line_login.setPointSize(11)
font_line_login.setItalic(True)
line_login.setFont(font_line_login)
# дополнительные параметры
line_login.setEchoMode(QLineEdit.Normal)
# устанавливаем исчезающий текст
line_login.setPlaceholderText('login')
line_login.setText('admin')

# создаём однострочное поле для ввода пароля
line_password = QLineEdit(window)
# присваиваем уникальное объектное имя полю для ввода пароля
line_password.setObjectName(u"line_password")
# устанавливаем положение и размер поля для ввода пароля в родительском окне (window)
line_password.setGeometry(QRect(1470, 60, 111, 31))
# задаём параметры стиля и оформления поля для ввода пароля
font_line_password = QFont()
font_line_password.setFamily(u"Arial")
font_line_password.setPointSize(11)
font_line_password.setItalic(True)
line_password.setFont(font_line_password)
# дополнительные параметры
line_password.setEchoMode(QLineEdit.Password)
# устанавливаем исчезающий текст
line_password.setPlaceholderText('password')
line_password.setText('admin')

# устанавливаем надпись "Логин"
label_login = QLabel('Логин', window)
# присваиваем уникальное объектное имя надписи "Логин"
label_login.setObjectName(u"label_login")
# устанавливаем положение и размер поля для надписи "Логин" в родительском окне (window)
label_login.setGeometry(QRect(1400, 30, 61, 21))
# задаём параметры стиля и оформления поля для надписи "Логин"
font_label_login = QFont()
font_label_login.setFamily(u"Arial")
font_label_login.setPointSize(12)
font_label_login.setItalic(True)
label_login.setFont(font_label_login)

# устанавливаем надпись "Пароль"
label_password = QLabel('Пароль', window)
# присваиваем уникальное объектное имя надписи "Пароль"
label_password.setObjectName(u"label_password")
# устанавливаем положение и размер поля для надписи "Пароль" в родительском окне (window)
label_password.setGeometry(QRect(1390, 70, 81, 21))
# задаём параметры стиля и оформления поля для надписи "Пароль"
font_label_password = QFont()
font_label_password.setFamily(u"Arial")
font_label_password.setPointSize(12)
font_label_password.setItalic(True)
# скрываем введённые с клавиатуры символы при вводе в поле "Пароль"
label_password.setFont(font_label_password)

# устанавливаем надпись "Линия"
label_line = QLabel('Номер линии', window)
# присваиваем уникальное объектное имя надписи "Линия"
label_line.setObjectName(u"label_line")
# устанавливаем положение и размер поля для надписи "Линия" в родительском окне (window)
label_line.setGeometry(QRect(20, 25, 151, 21))
# задаём параметры стиля и оформления поля для надписи "Линия"
font_label_line = QFont()
font_label_line.setFamily(u"Arial")
font_label_line.setPointSize(12)
font_label_line.setItalic(True)
label_line.setFont(font_label_line)
label_line.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Чертёж"
label_drawing = QLabel('Номер чертежа', window)
# присваиваем уникальное объектное имя надписи "Чертёж"
label_drawing.setObjectName(u"label_drawing")
# устанавливаем положение и размер поля для надписи "Чертёж" в родительском окне (window)
label_drawing.setGeometry(QRect(20, 65, 151, 21))
# задаём параметры стиля и оформления поля для надписи "Чертёж"
font_label_drawing = QFont()
font_label_drawing.setFamily(u"Arial")
font_label_drawing.setPointSize(12)
font_label_drawing.setItalic(True)
label_drawing.setFont(font_label_drawing)
label_drawing.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Юнит"
label_unit = QLabel('Юнит', window)
# присваиваем уникальное объектное имя надписи "Юнит"
label_unit.setObjectName(u"label_label_unit")
# устанавливаем положение и размер поля для надписи "Юнит" в родительском окне (window)
label_unit.setGeometry(QRect(20, 105, 151, 21))
# задаём параметры стиля и оформления поля для надписи "Юнит"
font_label_unit = QFont()
font_label_unit.setFamily(u"Arial")
font_label_unit.setPointSize(12)
font_label_unit.setItalic(True)
label_unit.setFont(font_label_unit)
label_unit.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Номер локации"
label_item_description = QLabel('Номер локации', window)
# присваиваем уникальное объектное имя надписи "Номер локации"
label_item_description.setObjectName(u"label_item_description")
# устанавливаем положение и размер поля для надписи "Номер локации" в родительском окне (window)
label_item_description.setGeometry(QRect(360, 105, 151, 21))
# задаём параметры стиля и оформления поля для надписи "Номер локации"
font_label_item_description = QFont()
font_label_item_description.setFamily(u"Arial")
font_label_item_description.setPointSize(12)
font_label_item_description.setItalic(True)
label_item_description.setFont(font_label_item_description)
label_item_description.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Номер отчёта"
label_number_report = QLabel('Номер отчёта', window)
# присваиваем уникальное объектное имя надписи "Номер локации"
label_number_report.setObjectName(u"label_number_report")
# устанавливаем положение и размер поля для надписи "Номер локации" в родительском окне (window)
label_number_report.setGeometry(QRect(20, 145, 151, 21))
# задаём параметры стиля и оформления поля для надписи "Номер локации"
font_label_number_report = QFont()
font_label_number_report.setFamily(u"Arial")
font_label_number_report.setPointSize(12)
font_label_number_report.setItalic(True)
label_number_report.setFont(font_label_number_report)
label_number_report.setAlignment(Qt.AlignRight)

# создаём кнопку "Войти"
button_log_in = QPushButton('Войти', window)
# устанавливаем положение и размер кнопки "Войти" в родительском окне (window)
button_log_in.setGeometry(QRect(1590, 20, 111, 31))
# присваиваем уникальное объектное имя кнопке "Войти"
button_log_in.setObjectName(u"pushButton_enter")
# задаём параметры стиля и оформления кнопки "Войти"
font_button_log_in = QFont()
font_button_log_in.setFamily(u"Arial")
font_button_log_in.setPointSize(14)
button_log_in.setFont(font_button_log_in)
# дополнительные параметры
button_log_in.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку "Выйти"
button_log_out = QPushButton('Выйти', window)
# устанавливаем положение и размер кнопки "Выйти" в родительском окне (window)
button_log_out.setGeometry(QRect(1590, 60, 111, 31))
# присваиваем уникальное объектное имя кнопке "Выйти"
button_log_out.setObjectName(u"pushButton_out")
# задаём параметры стиля и оформления кнопки "Выйти"
font_button_log_out = QFont()
font_button_log_out.setFamily(u"Arial")
font_button_log_out.setPointSize(14)
button_log_out.setFont(font_button_log_out)
# дополнительные параметры
button_log_out.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Выйти" до авторизации
button_log_out.setDisabled(True)

# создаём кнопку "Удалить"
button_delete = QPushButton('Удалить', window)
# устанавливаем положение и размер кнопки "Удалить" в родительском окне (window)
button_delete.setGeometry(QRect(20, 904, 171, 41))
# присваиваем уникальное объектное имя кнопке "Удалить"
button_delete.setObjectName(u"pushButton_delete")
# задаём параметры стиля и оформления кнопки "Удалить"
font_button_delete = QFont()
font_button_delete.setFamily(u"Arial")
font_button_delete.setPointSize(14)
button_delete.setFont(font_button_delete)
# дополнительные параметры
button_delete.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Удалить" до авторизации
button_delete.setDisabled(True)

# создаём кнопку "Сводные данные"
button_statistic_master = QPushButton('Сводные данные', window)
# устанавливаем положение и размер кнопки "Сводные данные" в родительском окне (window)
button_statistic_master.setGeometry(QRect(761, 904, 200, 41))
# присваиваем уникальное объектное имя кнопке "Сводные данные"
button_statistic_master.setObjectName(u"pushButton_statistic_master")
# задаём параметры стиля и оформления кнопки "Сводные данные"
font_button_statistic_master = QFont()
font_button_statistic_master.setFamily(u"Arial")
font_button_statistic_master.setPointSize(14)
button_statistic_master.setFont(font_button_statistic_master)
# дополнительные параметры
button_statistic_master.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Сводные данные" до авторизации
button_statistic_master.setDisabled(True)

# вставляем картинку YKR
label_ykr = QLabel(window)
label_ykr.setObjectName(u"Rutledge")
label_ykr.setGeometry(QRect(990, 10, 111, 121))
label_ykr.setPixmap(QPixmap(u"logo_ykr.png"))

# вставляем картинку NCA
label_nca = QLabel(window)
label_nca.setObjectName(u"Rutledge")
label_nca.setGeometry(QRect(1120, 10, 111, 121))
label_nca.setPixmap(QPixmap(u"logo_nca.png"))

# вставляем картинку NCOC
label_ncoc = QLabel(window)
label_ncoc.setObjectName(u"Rutledge")
label_ncoc.setGeometry(QRect(1250, 13, 111, 115))
label_ncoc.setPixmap(QPixmap(u"logo_ncoc.png"))

# общая область с боковой полосой прокрутки
scroll_area = QScrollArea(window)
scroll_area.setObjectName(u'Scroll_Area')
# полоса прокрутки появляется, только если таблицы больше самой области прокрутки
scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
# задаём размер области с полосой прокрутки
scroll_area.setGeometry(20, 245, 1681, 650)

# создаём группу из чек-бокса 'ON', 'OF', 'OS'
groupBox_location = QGroupBox(window)
groupBox_location.setObjectName(u"groupBox_radio")
# устанавливаем размер группы чек-бокса
groupBox_location.setGeometry(QRect(20, 180, 161, 56))
# устанавливаем название группы чек-бокса
groupBox_location.setTitle('Локация')
groupBox_location.setStyleSheet('''QGroupBox {border: 0.5px solid grey;};
                                   QGroupBox:title{
                                   subcontrol-origin: margin;
                                   subcontrol-position: top center;
                                   padding: 0 3px 0 3px;
                                }''')

# создаём радио-кнопку локации 'ON'
checkBox_on = QCheckBox(groupBox_location)
checkBox_on.setObjectName(u"checkBox_on")
# устанавливаем положение внутри группы
checkBox_on.setGeometry(QRect(10, 25, 42, 20))
# указываем текст чек-бокса
checkBox_on.setText('ON')
checkBox_on.setChecked(True)

# создаём радио-кнопку локации 'OF'
checkBox_of = QCheckBox(groupBox_location)
checkBox_of.setObjectName(u"checkBox_of")
# устанавливаем положение внутри группы
checkBox_of.setGeometry(QRect(60, 25, 42, 20))
# указываем текст чек-бокса
checkBox_of.setText('OF')

# создаём радио-кнопку локации 'OS'
checkBox_os = QCheckBox(groupBox_location)
checkBox_os.setObjectName(u"checkBox_os")
# устанавливаем положение внутри группы
checkBox_os.setGeometry(QRect(110, 25, 42, 20))
# указываем текст чек-бокса
checkBox_os.setText('OS')

# создаём группу из чек-боксов методов контроля
groupBox_ndt = QGroupBox(window)
groupBox_ndt.setObjectName(u"groupBox_ndt")
groupBox_ndt.setGeometry(QRect(190, 180, 161, 56))
# устанавливаем название группы чек-боксов
groupBox_ndt.setTitle('Метод контроля')
groupBox_ndt.setStyleSheet('''QGroupBox {border: 0.5px solid grey;};
                              QGroupBox:title{
                              subcontrol-origin: margin;
                              subcontrol-position: top center;
                              padding: 0 3px 0 3px;
                           }''')

# создаём чек-бокс метода контроля 'UTT'
checkBox_utt = QCheckBox(groupBox_ndt)
checkBox_utt.setObjectName(u"checkBox_utt")
# устанавливаем положение внутри группы
checkBox_utt.setGeometry(QRect(10, 25, 61, 20))
# указываем текст чек-бокса
checkBox_utt.setText('UTT')
# делаем чек-бокс 'UTT' активным по умолчанию
checkBox_utt.setChecked(True)

# создаём чек-бокс метода контроля 'PAUT'
checkBox_paut = QCheckBox(groupBox_ndt)
checkBox_paut.setObjectName(u"checkBox_paut")
# устанавливаем положение внутри группы
checkBox_paut.setGeometry(QRect(80, 25, 61, 20))
# указываем текст чек-бокса
checkBox_paut.setText('PAUT')

# создаём группу из чек-боксов годов
groupBox_year = QGroupBox(window)
groupBox_year.setObjectName(u"groupBox_year")
# устанавливаем размер группы радио-кнопок
groupBox_year.setGeometry(QRect(360, 180, 381, 56))
# устанавливаем название группы чек-боксов
groupBox_year.setTitle('Год контроля')
groupBox_year.setStyleSheet('''QGroupBox {border: 0.5px solid grey;};
                               QGroupBox:title{
                               subcontrol-origin: margin;
                               subcontrol-position: top center;
                               padding: 0 3px 0 3px;
                            }''')

# создаём чек-бокс года '2023'
checkBox_2023 = QCheckBox(groupBox_year)
checkBox_2023.setObjectName(u"checkBox_2023")
# устанавливаем положение внутри группы
checkBox_2023.setGeometry(QRect(10, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2023.setText('2023')

# создаём чек-бокс года '2022'
checkBox_2022 = QCheckBox(groupBox_year)
checkBox_2022.setObjectName(u"checkBox_2022")
# устанавливаем положение внутри группы
checkBox_2022.setGeometry(QRect(80, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2022.setText('2022')
# делаем чек-бокс '2022' активным по умолчанию
checkBox_2022.setChecked(True)

# создаём чек-бокс года '2021'
checkBox_2021 = QCheckBox(groupBox_year)
checkBox_2021.setObjectName(u"checkBox_2021")
# устанавливаем положение внутри группы
checkBox_2021.setGeometry(QRect(150, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2021.setText('2021')

# создаём чек-бокс года '2020'
checkBox_2020 = QCheckBox(groupBox_year)
checkBox_2020.setObjectName(u"checkBox_2020")
# устанавливаем положение внутри группы
checkBox_2020.setGeometry(QRect(220, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2020.setText('2020')

# создаём чек-бокс года '2019'
checkBox_2019 = QCheckBox(groupBox_year)
checkBox_2019.setObjectName(u"checkBox_2019")
# устанавливаем положение внутри группы
checkBox_2019.setGeometry(QRect(290, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2019.setText('2019')


# нажатие кнопки "Войти"
def log_in():
    # если ничего не введено в поля "Логин" и "Пароль"
    if line_login.text() == '' and line_password.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы ничего не ввели в поля "Логин" и "Пароль"!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error('Попытка авторизоваться - не заполнены поля "Логин" и "Пароль"')
    # если ничего не введено в поле "Логин"
    elif line_login.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не заполнили поле "Логин"!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error(
            'Попытка авторизоваться - не заполнено поле "Логин", указан пароль - "{}"'.format(line_password.text()))
    # если ничего не введено в поле "Пароль"
    elif line_password.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не заполнили поле "Пароль"!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error(
            'Попытка авторизоваться - Не заполнено поле "Пароль", указан логин - "{}"'.format(line_login.text()))
    # если правильно введён логин и пароль
    elif line_login.text() == 'admin' and line_password.text() == 'admin':
        # делаем активными кнопки "Добавить", "Удалить", "Выйти", "Сводные данные"
        button_delete.setDisabled(False)
        button_log_out.setDisabled(False)
        button_add.setDisabled(False)
        button_statistic_master.setDisabled(False)
        # очищаем поле ввода логина и пароля
        line_login.clear()
        line_password.clear()
        # блокируем кнопку "Войти"
        button_log_in.setDisabled(True)
        logger_with_user.info('Пользователь авторизовался')
        # # делаем видимые флажки
        # if list_check_box:
        #     for i in list_check_box:
        #         i.show()
        # # сигнал о том, что выполнена авторизация
        # global authorization
        # authorization += 1
    # если неправильно введён логин или пароль
    else:
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы ввели не правильный логин или пароль!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error(
            'Попытка авторизоваться - Введён неверный логин "{}" или пароль "{}"'.format(line_login.text(), line_password.text()))


# нажатие кнопки "Выйти"
def log_out():
    # делаем НЕ активными кнопки "Добавить", "Редактировать", "Удалить", "Выйти", "Сводные данные", "Выйти"
    button_delete.setDisabled(True)
    button_log_out.setDisabled(True)
    button_add.setDisabled(True)
    button_statistic_master.setDisabled(True)
    # разблокируем кнопку "Войти"
    button_log_in.setDisabled(False)
    logger_with_user.info('Пользователь вышел')
    # # сбрасываем на ноль авторизацию для отображения флажков
    # global authorization
    # authorization = 0
    # # если перед выходом из авторизации показана НЕ статистика из master (check_statistic_master == 0),
    # if check_statistic_master == 0:
    #     # то скрываем флажки
    #     if list_check_box:
    #         for i in list_check_box:
    #             i.hide()


# словарь выбранных фильтров (локация, метод, год) для поиска
data_filter_for_search = dict()
data_filter_for_search['location'] = {'ON': True, 'OF': False, 'OS': False}
data_filter_for_search['method'] = {'UTT': True, 'PAUT': False}
data_filter_for_search['year'] = {'2019': False, '2020': False, '2021': False, '2022': True, '2023': False}


# обработчик события выбора одной из локаций (on, of, os)
def on_button_clicked_location():
    check_button_location = QObject().sender()
    data_filter_for_search['location'][check_button_location.text()] = check_button_location.isChecked()


# определяем какая локация выбрана
for button in groupBox_location.findChildren(QCheckBox):
    button.clicked.connect(on_button_clicked_location)


# обработчик события выбора одного из методов контроля (utt, paut)
def on_button_clicked_ndt():
    check_button_ndt = QObject().sender()
    data_filter_for_search['method'][check_button_ndt.text()] = check_button_ndt.isChecked()


# определяем какой(-ие) методы контроля выбраны
for button in groupBox_ndt.findChildren(QCheckBox):
    button.clicked.connect(on_button_clicked_ndt)


# обработчик события выбора одного из фильтров (локация, метод, год)
def on_button_clicked_year():
    check_button_year = QObject().sender()
    data_filter_for_search['year'][check_button_year.text()] = check_button_year.isChecked()


# определяем какой(-ие) года выбраны
for button in groupBox_year.findChildren(QCheckBox):
    button.clicked.connect(on_button_clicked_year)


# нажатие на кнопку "Поиск"
def search():
    # словарь введённых данных в поля для поиска
    data_for_search = dict()
    # определяем какие данные для поиска введены в поля для поиска
    data_for_search['line_search'] = [line_search_line.text()]
    data_for_search['drawing_search'] = [line_search_drawing.text()]
    data_for_search['unit'] = [line_search_unit.text()]
    data_for_search['item_description_search'] = [line_search_item_description.text()]
    data_for_search['number_report_search'] = [line_search_number_report.text()]
    # проверка - введены ли данные для поиска и выбраны ли все фильтры
    if not_check_data_and_filter(data_filter_for_search, data_for_search):
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не ввели данные для поиска или не выбрали не один фильтр',
            buttons=QMessageBox.Ok
        )
        return
    # собираем названия БД, по выбранным фильтрам, в которых надо искать данные
    db_for_search = define_db_for_search(data_filter_for_search)
    # получаем значения из полей для ввода
    values_for_search = dict()
    # путь поиска, в зависимости от количества заполненных полей для поиска
    search_path = 0
    values_for_search['line'] = line_search_line.text()
    if line_search_line.text():
        search_path += 1
    values_for_search['drawing'] = line_search_drawing.text()
    if line_search_drawing.text():
        search_path += 1
    values_for_search['unit'] = line_search_unit.text()
    if line_search_unit.text():
        search_path += 1
    values_for_search['item_description'] = line_search_item_description.text()
    if line_search_item_description.text():
        search_path += 1
    values_for_search['number_report'] = line_search_number_report.text()
    if line_search_number_report.text():
        search_path += 1
    print(db_for_search)
    print(values_for_search)
    print(search_path)
    # ищем данные в БД
    table_for_view = look_up_data(db_for_search, values_for_search, search_path)



    # !!!!!!!!! создание фрейма и помещение его в область для вывода сделать отдельной функцией
    # !!!!!!!!! расчёты размеров для отображения найденных данных отельной функцией

    # frame в который будут вставляться, таблицы чтобы при большом количестве таблиц появлялась полоса прокрутки
    frame_for_table = QFrame()
    # помещаем frame в область с полосой прокрутки
    scroll_area.setWidget(frame_for_table)
    # задаём размер frame
    frame_for_table.setGeometry(0, 0, 1681, 200)
    frame_for_table.show()
    # задаём поле для вывода данных из базы данных, размещённую в области с полосой прокрутки
    table_view = QTableView(frame_for_table)
    # создаём модель
    sqm = QSqlQueryModel(parent=window)
    # устанавливаем ширину столбцов под содержимое
    table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    # устанавливаем высоту столбцов под содержимое
    table_view.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    # устанавливаем разный цвет фона для чётных и нечётных строк
    table_view.setAlternatingRowColors(True)
    table_view.setModel(sqm)


    
    # выводим данные в форму из найденных таблиц
    # sqm.setQuery('SELECT * FROM {} WHERE line LIKE "%{}%"'.format(table_for_search_line[i], line_for_search))



# нажатие кнопки "Войти"
button_log_in.clicked.connect(log_in)

# нажатие на кнопку "Выйти"
button_log_out.clicked.connect(log_out)

# нажатие на кнопку "Поиск"
button_search.clicked.connect(search)
# нажатие на кнопку Enter когда фокус (каретка - мигающий символ "|") находится в поле для ввода номера линии, чертежа, локации или номера репорта
line_search_line.returnPressed.connect(search)
line_search_drawing.returnPressed.connect(search)
line_search_item_description.returnPressed.connect(search)
line_search_number_report.returnPressed.connect(search)


def main():
    try:
        window.show()
        sys.exit(app.exec_())
    finally:
        logger_with_user.info('Программа закрыта\n'
                              '--------------------------------------------------------------------------------')


if __name__ == '__main__':
    main()
