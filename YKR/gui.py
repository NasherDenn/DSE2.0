from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtSql import QSqlDatabase
from PyQt5.QtSql import QSqlQueryModel
import sys
from back_end import *

# создаём приложение
app = QApplication(sys.argv)
# создаём окно приложения
window = QWidget()
# название приложения
window.setWindowTitle('Finder')
# задаём стиль приложения Fusion
app.setStyle('Fusion')
# размер окна приложения
window.setFixedSize(1524, 872)

# устанавливаем favicon в окне приложения
icon = QIcon()
icon.addFile(u"favicon.png", QSize(), QIcon.Normal, QIcon.Off)
icon.addFile(u"favicon.png", QSize(), QIcon.Active, QIcon.On)
app.setWindowIcon(icon)

# создаём однострочное поле для ввода номер линии или чертежа
line_search = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search.setGeometry(QRect(20, 20, 561, 41))
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search.setObjectName(u"line_search")
# задаём параметры стиля и оформления окна ввода
font_line_search = QFont()
font_line_search.setFamily(u"Arial")
font_line_search.setPointSize(14)
font_line_search.setItalic(False)
# дополнительные параметры
line_search.setFont(font_line_search)
line_search.setMouseTracking(False)
line_search.setFocusPolicy(Qt.ClickFocus)
line_search.setContextMenuPolicy(Qt.NoContextMenu)
line_search.setAcceptDrops(True)
line_search.setStyleSheet(u"")
line_search.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search.setEchoMode(QLineEdit.Normal)
line_search.setCursorPosition(0)
line_search.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search.setClearButtonEnabled(True)
line_search.setText('28235724')
line_search.setFocus()

# создаём кнопку "Поиск"
button_search = QPushButton('Поиск', window)
# устанавливаем положение и размер кнопки для поиска в родительском окне (window)
button_search.setGeometry(600, 20, 161, 41)
# присваиваем уникальное объектное имя кнопке "Поиск"
button_search.setObjectName(u"pushButton_search")
# задаём параметры стиля и оформления кнопки "Поиск"
font_button_search = QFont()
font_button_search.setFamily(u"Arial")
font_button_search.setPointSize(14)
button_search.setFont(font_button_search)
# дополнительные параметры
button_search.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку "Закрыть" из программы
button_exit = QPushButton('Закрыть', window)
# устанавливаем положение и размер кнопки "Закрыть" для выхода из программы в родительском окне (window)
button_exit.setGeometry(QRect(1340, 810, 161, 41))
# присваиваем уникальное объектное имя кнопке "Закрыть"
button_exit.setObjectName(u"pushButton_exit")
# задаём параметры стиля и оформления кнопки "Закрыть"
font_button_exit = QFont()
font_button_exit.setFamily(u"Arial")
font_button_exit.setPointSize(14)
button_exit.setFont(font_button_exit)
# дополнительные параметры
button_exit.setFocusPolicy(Qt.ClickFocus)
# Закрыть из программы при нажатии на кнопку "Закрыть"
button_exit.clicked.connect(qApp.exit)

# создаём однострочное поле для ввода логина
line_login = QLineEdit(window)
# присваиваем уникальное объектное имя полю для ввода логина
line_login.setObjectName(u"line_login")
# устанавливаем положение и размер поля для ввода логина в родительском окне (window)
line_login.setGeometry(QRect(1270, 20, 111, 31))
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

# создаём однострочное поле для ввода пароля
line_password = QLineEdit(window)
# присваиваем уникальное объектное имя полю для ввода пароля
line_password.setObjectName(u"line_password")
# устанавливаем положение и размер поля для ввода пароля в родительском окне (window)
line_password.setGeometry(QRect(1270, 60, 111, 31))
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

# устанавливаем надпись "Логин"
label_login = QLabel('Логин', window)
# присваиваем уникальное объектное имя надписи "Логин"
label_login.setObjectName(u"label_login")
# устанавливаем положение и размер поля для надписи "Логин" в родительском окне (window)
label_login.setGeometry(QRect(1200, 30, 61, 21))
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
label_password.setGeometry(QRect(1190, 70, 81, 21))
# задаём параметры стиля и оформления поля для надписи "Пароль"
font_label_password = QFont()
font_label_password.setFamily(u"Arial")
font_label_password.setPointSize(12)
font_label_password.setItalic(True)
# скрываем введённые с клавиатуры символы при вводе в поле "Пароль"
label_password.setFont(font_label_password)

# создаём кнопку печати
button_print = QPushButton('Печать', window)
# устанавливаем положение и размер кнопки печати в родительском окне (window)
button_print.setGeometry(QRect(20, 80, 161, 41))
# присваиваем уникальное объектное имя кнопке "Печать"
button_print.setObjectName(u"pushButton_print")
# задаём параметры стиля и оформления кнопки печати
font_button_print = QFont()
font_button_print.setFamily(u"Arial")
font_button_print.setPointSize(14)
button_print.setFont(font_button_print)
# дополнительные параметры
button_print.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку "Войти"
button_log_in = QPushButton('Войти', window)
# устанавливаем положение и размер кнопки "Войти" в родительском окне (window)
button_log_in.setGeometry(QRect(1390, 20, 111, 31))
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
button_log_out.setGeometry(QRect(1390, 60, 111, 31))
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

# создаём кнопку "Добавить"
button_add = QPushButton('Добавить', window)
# устанавливаем положение и размер кнопки "Добавить" в родительском окне (window)
button_add.setGeometry(QRect(200, 80, 161, 41))
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
# button_add.setDisabled(True)
button_add.setDisabled(False)

# создаём кнопку "Редактировать"
button_repair = QPushButton('Редактировать', window)
# устанавливаем положение и размер кнопки "Редактировать" в родительском окне (window)
button_repair.setGeometry(QRect(380, 80, 201, 41))
# присваиваем уникальное объектное имя кнопке "Редактировать"
button_repair.setObjectName(u"pushButton_repair")
# задаём параметры стиля и оформления кнопки "Редактировать"
font_button_repair = QFont()
font_button_repair.setFamily(u"Arial")
font_button_repair.setPointSize(14)
button_repair.setFont(font_button_repair)
# дополнительные параметры
button_repair.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Редактировать" до авторизации
button_repair.setDisabled(True)

# создаём кнопку "Готово"
button_ok = QPushButton('Готово', window)
# устанавливаем положение и размер кнопки "Готово" в родительском окне (window)
button_ok.setGeometry(QRect(600, 80, 161, 41))
# присваиваем уникальное объектное имя кнопке "Готово"
button_ok.setObjectName(u"pushButton_ok")
# задаём параметры стиля и оформления кнопки "Готово"
font_button_ok = QFont()
font_button_ok.setFamily(u"Arial")
font_button_ok.setPointSize(14)
button_ok.setFont(font_button_ok)
# дополнительные параметры
button_ok.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку "Удалить"
button_delete = QPushButton('Удалить', window)
# устанавливаем положение и размер кнопки "Удалить" в родительском окне (window)
button_delete.setGeometry(QRect(20, 810, 171, 41))
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

# вставляем картинку YKR
label_ykr = QLabel(window)
label_ykr.setObjectName(u"Rutledge")
label_ykr.setGeometry(QRect(790, 10, 111, 121))
label_ykr.setPixmap(QPixmap(u"logo_ykr.png"))

# вставляем картинку NCA
label_nca = QLabel(window)
label_nca.setObjectName(u"Rutledge")
label_nca.setGeometry(QRect(920, 10, 111, 121))
label_nca.setPixmap(QPixmap(u"logo_nca.png"))

# вставляем картинку NCOC
label_ncoc = QLabel(window)
label_ncoc.setObjectName(u"Rutledge")
label_ncoc.setGeometry(QRect(1050, 13, 111, 115))
label_ncoc.setPixmap(QPixmap(u"logo_ncoc.png"))

# общая область с боковой полосой прокрутки
scroll_area = QScrollArea(window)
scroll_area.setObjectName(u'Scroll_Area')
# полоса прокрутки появляется, только если таблицы больше самой области прокрутки
scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
# задаём размер области с полосой прокрутки
scroll_area.setGeometry(20, 140, 1481, 650)


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
    # если ничего не введено в поле "Логин"
    elif line_login.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не заполнили поле "Логин"!!!',
            buttons=QMessageBox.Ok
        )
    # если ничего не введено в поле "Пароль"
    elif line_password.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не заполнили поле "Пароль"!!!',
            buttons=QMessageBox.Ok
        )
    # если правильно введён логин и пароль
    elif line_login.text() == 'admin' and line_password.text() == 'admin':
        # делаем активными кнопки "Добавить", "Редактировать", "Готово", "Удалить", "Выйти"
        button_repair.setDisabled(False)
        button_delete.setDisabled(False)
        button_log_out.setDisabled(False)
        # очищаем поле ввода логина и пароля
        line_login.clear()
        line_password.clear()
        # блокируем кнопку "Войти"
        button_log_in.setDisabled(True)
    # если неправильно введён логин или пароль
    else:
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы ввели не правильный логин или пароль!!!',
            buttons=QMessageBox.Ok
        )


# нажатие на кнопку "Добавить"
def add_tables():
    name_dir = QFileDialog.getExistingDirectory(None, 'Выбрать папку', '.')
    add_table(name_dir)


# нажатие на кнопку "Поиск"
def search():
    if line_search.text():
        # удаляем обозначения дюймов "
        if re.findall(r'\'\'|"|”', line_search.text()):
            line_for_search = re.sub(r'"|\'\'', '', line_search.text()).upper()
        # получаем текст из поля для ввода и приводим его в верхний регистр
        line_for_search = line_search.text().upper()
        # создаём соединение с базой данной
        con = QSqlDatabase.addDatabase('QSQLITE')
        # передаём имя базы данных для открытия
        con.setDatabaseName(r'C:\Users\asus\PycharmProjects\YKR\YKR\reports_db.sqlite')
        # если соединение не установлено, то сообщение об ошибке и выход
        if not con.open():
            QMessageBox.critical(
                None,
                'App name Error',
                'Error to connect to the database')
            sys.exit()
        else:
            # список таблиц в которой есть искомая линия
            table_for_search_line = []
            # список таблиц в которой есть искомый чертёж
            table_for_search_drawing = []
            # список таблиц в которой есть искомый номер репорта
            table_for_search_report = []
            # список таблиц в которой есть искомый номер work order
            table_for_search_wo = []
            # список репортов из sqlite_master где есть номер work order
            reports_for_search_wo = []
            # проверяем наличие областей tableView для вывода данных
            # если есть, то закрываем их, чтобы не наслаивались
            if window.findChildren(QTableView):
                open_tableview = window.findChildren(QTableView)
                for i in open_tableview:
                    i.hide()

            # перебираем таблицы, которые попали в базу данных после очистки
            for i in con.tables():
                # подключаемся в базе данных
                conn = sqlite3.connect('reports_db.sqlite')
                cur = conn.cursor()
                # перебираем список названий столбцов в таблице
                for k in cur.execute('SELECT * FROM {}'.format(i)).description:
                    # если 'Line' есть в названии столбца
                    if 'Line' in k:
                        # и если искомая линия есть в таблице, то добавляем имя таблицы в список table_for_search_line
                        if cur.execute('SELECT Line FROM {} WHERE Line="{}"'.format(i, line_for_search)).fetchall():
                            table_for_search_line.append(i)
                    # если 'Drawing' есть в названии столбца
                    if 'Drawing' in k:
                        # и если искомый чертёж есть в таблице, то добавляем имя таблицы в список
                        # table_for_search_drawing
                        if cur.execute(
                                'SELECT Drawing FROM {} WHERE Drawing="{}"'.format(i, line_for_search)).fetchall():
                            table_for_search_drawing.append(i)
                cur.close()

            # если в поле для поиска указан номер репорта
            if line_for_search[:6] == '04-YKR':
                # перебираем sqlite_master в поиске репорта
                # подключаемся в базе данных
                conn = sqlite3.connect('reports_db.sqlite')
                cur = conn.cursor()
                # меняем '-' на '_'
                line_for_search_report = re.sub('-', '_', line_for_search)
                # если нашли номер репорта в sqlite_master
                if cur.execute(
                        'SELECT tbl_name FROM sqlite_master WHERE name LIKE "%{}"'.format(line_for_search_report)):
                    table_for_search_report.append(cur.execute(
                        'SELECT tbl_name FROM sqlite_master WHERE name LIKE "%{}"'.format(
                            line_for_search_report)).fetchall())
                cur.close()

            # если в поле для поиска указан номер work order
            if line_for_search.isdigit():
                # перебираем master в поиске work order
                # подключаемся в базе данных
                conn = sqlite3.connect('reports_db.sqlite')
                cur = conn.cursor()
                # если нашли work order в master
                if cur.execute(
                        'SELECT report_number FROM master WHERE work_order="{}"'.format(line_for_search)):
                    reports_for_search_wo.append(cur.execute(
                        'SELECT report_number FROM master WHERE work_order="{}"'.format(line_for_search)).fetchall())
                    # перебираем найденные номера репортов
                    for ii in reports_for_search_wo[0]:
                        # добавляем в список если нашли номера таблиц
                        table_for_search_wo.append(cur.execute(
                            'SELECT tbl_name FROM sqlite_master WHERE name LIKE "%{}"'.format(ii[0])).fetchall())
                # закрываем соединение
                cur.close()

            # если для поиска указан не номер репорта или work order
            def depthCount(x, depth=0):
                if not x or not isinstance(x, list):
                    return depth
                return max(depthCount(x[0], depth + 1),
                           depthCount(x[1:], depth))

            # если для поиска указан не номер репорта
            if depthCount(table_for_search_report) == 0:
                table_for_search_report = [[]]

            # если найден номер линии или номер чертежа, или номер репорта, или номер work order, то показываем область
            # для таблицы с найденными данными
            if table_for_search_line or table_for_search_drawing or table_for_search_report or table_for_search_wo:
                if len(table_for_search_line) + len(table_for_search_drawing) + len(
                        table_for_search_report[0]) + len(table_for_search_wo) != 0:
                    # считаем количество найденных таблиц для вывода нужного количества tableView
                    count_table_view = len(table_for_search_line) + len(table_for_search_drawing) + len(
                        table_for_search_report[0]) + len(table_for_search_wo)
                    # список названий таблицы для переменной при создании tableView
                    table_view = ['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
                                  'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen',
                                  'eighteen',
                                  'nineteen', 'twenty', 'twenty_one', 'twenty_two', 'twenty_three', 'twenty_four',
                                  'twenty_five', 'twenty_six', 'twenty_seven', 'twenty_eight', 'twenty_nine', 'thirty']
                    # frame в который будут вставляться, таблицы чтобы при большом количестве таблиц появлялась полоса
                    # прокрутки
                    frame_for_table = QFrame()
                    # подключаемся в базе данных
                    cur = conn.cursor()
                    # список количества строк в каждой найденной таблице
                    count_row_table_view = []
                    if table_for_search_line:
                        for i in table_for_search_line:
                            # количество строк в одной найденной таблице count_row_table_view[0][0]
                            count_row_table = cur.execute(
                                'SELECT COUNT(*) FROM {} WHERE Line="{}"'.format(i, line_for_search)).fetchall()
                            count_row_table_view.append(count_row_table[0][0])
                    if table_for_search_drawing:
                        for i in table_for_search_drawing:
                            # количество строк в одной найденной таблице count_row_table[0][0]
                            count_row_table = cur.execute(
                                'SELECT COUNT(*) FROM {} WHERE Drawing="{}"'.format(i, line_for_search)).fetchall()
                            count_row_table_view.append(count_row_table[0][0])
                    if table_for_search_report:
                        for i in table_for_search_report[0]:
                            count_row_table = cur.execute('SELECT COUNT(*) FROM {}'.format(i[0])).fetchall()
                            count_row_table_view.append(count_row_table[0][0])
                    if table_for_search_wo:
                        for i in table_for_search_wo:
                            count_row_table = cur.execute('SELECT COUNT(*) FROM {}'.format(i[0][0])).fetchall()
                            count_row_table_view.append(count_row_table[0][0])
                    # закрываем соединение
                    cur.close()
                    # общее количество строк в найденных таблицах для длины frame
                    sum_row_table = 0
                    for i in count_row_table_view:
                        sum_row_table += i
                    # высота одной строки
                    one_row = 25
                    # высота фрейма = общее количество строк в найденных таблицах * высоту одной строки +
                    # + количество таблиц * 2 (кнопка номера репорта и строка названий столбцов) * 20 (высота одной
                    # строки) + 20 (высота первой строки с номером первого репорта) + количество таблиц * 20 (расстояние
                    # между таблицами в открытом виде
                    w = sum_row_table * one_row + len(count_row_table_view) * 2 * 20 + 20 + len(
                        count_row_table_view) * 20
                    # помещаем frame в область с полосой прокрутки
                    scroll_area.setWidget(frame_for_table)
                    # задаём размер frame
                    frame_for_table.setGeometry(0, 0, 1460, w)
                    frame_for_table.show()
                    # начальная координата y1 - первой кнопки с номером репорта первой, y2 - первой таблицы
                    y1 = 0
                    # список всех таблиц, номеров репортов и высоты каждой таблицы
                    list_table_view = []
                    list_button_for_table = []
                    list_height_table_view = []
                    # вытягиваем данные из найденных таблиц, формируем таблицу, кнопку названия репорта
                    for i in range(count_table_view):
                        # высота одной таблицы tableView = количество строк в одной таблице * высоту одной строки +
                        # + высота строки названия столбцов
                        height = count_row_table_view[i] * one_row + one_row
                        # создаём переменную названия кнопок номеров репортов для вывода данных
                        if table_for_search_line:
                            button_for_table = table_for_search_line[i]
                            # переменная для поиска даты и work order репорта в таблице master
                            for_w_o = button_for_table[(button_for_table.index('_04') + 1):]
                            # подключаемся в базе данных
                            cur = conn.cursor()
                            # переменная номера work order
                            w_o = cur.execute(
                                'SELECT report_date, work_order FROM master WHERE report_number="{}"'.format(
                                    for_w_o)).fetchall()
                            # закрываем соединение
                            cur.close()
                            button_for_table = re.sub(r'_', '-', button_for_table)
                            ind = button_for_table.index('-04') + 1
                            # название кнопки по номеру репорта
                            second_underlining = button_for_table[ind:]
                            # добавляем к названию кнопки дату и work order
                            second_underlining = second_underlining + '     Date: ' + w_o[0][0] + '     WO: ' + w_o[0][
                                1]
                        if table_for_search_drawing:
                            button_for_table = table_for_search_drawing[i]
                            # переменная для поиска даты и work order репорта в таблице master
                            for_w_o = button_for_table[(button_for_table.index('_04') + 1):]
                            # подключаемся в базе данных
                            cur = conn.cursor()
                            # переменная номера work order
                            w_o = cur.execute(
                                'SELECT report_date, work_order FROM master WHERE report_number="{}"'.format(
                                    for_w_o)).fetchall()
                            # закрываем соединение
                            cur.close()
                            button_for_table = re.sub(r'_', '-', button_for_table)
                            ind = button_for_table.index('-04') + 1
                            # название кнопки по номеру репорта
                            second_underlining = button_for_table[ind:]
                            # добавляем к названию кнопки дату и work order
                            second_underlining = second_underlining + '     Date: ' + w_o[0][0] + '     WO: ' + w_o[0][
                                1]
                        # определяем глубину вложенности списка заданного для поиска репорта
                        if depthCount(table_for_search_report) == 1:
                            table_for_search_report = []
                        if table_for_search_report:
                            button_for_table = table_for_search_report[0][i][0]
                            # переменная для поиска даты и work order репорта в таблице master
                            for_w_o = button_for_table[(button_for_table.index('_04') + 1):]
                            # подключаемся в базе данных
                            cur = conn.cursor()
                            # переменная номера work order
                            w_o = cur.execute(
                                'SELECT report_date, work_order FROM master WHERE report_number="{}"'.format(
                                    for_w_o)).fetchall()
                            # закрываем соединение
                            cur.close()
                            button_for_table = re.sub(r'_', '-', button_for_table)
                            ind = button_for_table.index('-04') + 1
                            # название кнопки по номеру репорта
                            second_underlining = button_for_table[ind:]
                            # добавляем к названию кнопки дату и work order
                            second_underlining = second_underlining + '     Date: ' + w_o[0][0] + '     WO: ' + w_o[0][
                                1]
                        # определяем глубину вложенности списка заданного для поиска work order
                        if depthCount(table_for_search_wo) == 1:
                            table_for_search_wo = []
                        if table_for_search_wo:
                            button_for_table = table_for_search_wo[i][0][0]
                            # номер work order из строки для поиска
                            w_o = line_for_search
                            # подключаемся в базе данных
                            cur = conn.cursor()
                            # переменная даты репорта
                            date_report = cur.execute(
                                'SELECT report_date FROM master WHERE work_order="{}"'.format(w_o)).fetchall()
                            # закрываем соединение
                            cur.close()
                            button_for_table = re.sub(r'_', '-', button_for_table)
                            ind = button_for_table.index('-04') + 1
                            # название кнопки по номеру репорта
                            second_underlining = button_for_table[ind:]
                            # добавляем к названию кнопки дату и work order
                            second_underlining = second_underlining + '     Date: ' + date_report[i][
                                0] + '     WO: ' + w_o

                        # задаём название кнопки по номеру репорта и помещаем внутрь frame
                        button_for_table = QPushButton(second_underlining, frame_for_table)
                        # задаём размеры и место расположения кнопки во frame
                        button_for_table.setGeometry(QRect(0, y1, 500, 20))
                        # задаём стиль шрифта
                        font_button_for_table = QFont()
                        font_button_for_table.setFamily(u"Calibri")
                        font_button_for_table.setPointSize(10)
                        button_for_table.setStyleSheet('text-align: left; font: bold italic')
                        button_for_table.setFont(font_button_for_table)
                        button_for_table.show()
                        # скрываем границы кнопки
                        button_for_table.setFlat(True)
                        # делаем кнопку переключателем
                        button_for_table.setCheckable(True)
                        # задаём поле для вывода данных из базы данных, размещённую в области с полосой прокрутки
                        table_view[i] = QTableView(frame_for_table)
                        # устанавливаем координаты расположения таблиц в области с полосой прокрутки
                        list_button_for_table.append(button_for_table)
                        list_table_view.append(table_view[i])
                        list_height_table_view.append(height)
                        table_view[i].show()
                        # сдвигаем все последующие кнопки и таблицы
                        y1 += 20
                        # создаём модель
                        sqm = QSqlQueryModel(parent=window)
                        # устанавливаем ширину столбцов под содержимое
                        table_view[i].horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                        # устанавливаем высоту столбцов под содержимое
                        table_view[i].verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                        # устанавливаем разный цвет фона для чётных и нечётных строк
                        table_view[i].setAlternatingRowColors(True)
                        table_view[i].setModel(sqm)
                        # создаём запрос
                        # выводим данные в форму из найденных таблиц по номеру линии, чертежа или репорта
                        if len(table_for_search_line) > 0:
                            sqm.setQuery(
                                'SELECT * FROM {} WHERE Line="{}"'.format(table_for_search_line[i], line_for_search),
                                db=QSqlDatabase('reports_db.sqlite'))
                        # выводим данные в форму из найденных таблиц по номеру чертежа в таблице
                        if len(table_for_search_drawing) > 0:
                            sqm.setQuery(
                                'SELECT * FROM {} WHERE Drawing="{}"'.format(table_for_search_drawing[i],
                                                                             line_for_search),
                                db=QSqlDatabase('reports_db.sqlite'))
                        # выводим данные в форму из найденных таблиц по номеру репорта
                        if len(table_for_search_report) > 0:
                            if len(table_for_search_report[0]) > 0:
                                sqm.setQuery('SELECT * FROM {}'.format(table_for_search_report[0][i][0]),
                                             db=QSqlDatabase('reports_db.sqlite'))
                        # выводим данные в форму из найденных таблиц по номеру word order
                        if len(table_for_search_wo) > 0:
                            if len(table_for_search_wo[0]) > 0:
                                sqm.setQuery('SELECT * FROM {}'.format(table_for_search_wo[i][0][0]),
                                             db=QSqlDatabase('reports_db.sqlite'))
                        table_view[i].hide()
                        # обработка нажатия на кнопку с номером репорта в frame
                        button_for_table.clicked.connect(
                            lambda: visible_table_view(list_table_view, list_button_for_table,
                                                       list_height_table_view))
                        # активируем кнопку в левом верхнем углу таблицы для выделения всей таблицы
                        table_view[i].setCornerButtonEnabled(True)
                        # подключаемся в базе данных
                        conn = sqlite3.connect('reports_db.sqlite')
                        cur = conn.cursor()
                        # перебираем таблицы и извлекаем данные
                        if table_for_search_line:
                            reader = cur.execute("SELECT * FROM {}".format(table_for_search_line[i]))
                        elif table_for_search_drawing:
                            reader = cur.execute("SELECT * FROM {}".format(table_for_search_drawing[i]))
                        elif table_for_search_report:
                            reader = cur.execute("SELECT * FROM {}".format(table_for_search_report[0][i][0]))
                        elif table_for_search_wo:
                            reader = cur.execute("SELECT * FROM {}".format(table_for_search_wo[i][0][0]))
                        # получаем список названий столбцов
                        name_column = [x[0] for x in reader.description]
                        # получаем номер столбца с номинальной толщиной
                        print(name_column)
                        number_column_nominal_thickness = name_column.index('Nominal_thickness')

                        # # кортеж списков (всех строк) в каждой таблице
                        # string_table = reader.fetchall()
                        # # print(string_table)
                        # # список индексов минимальных значений в таблице
                        # list_index_minimum_thickness = []
                        # # срез - список значений для поиска минимального значения
                        # for j in range(len(string_table)):
                        #     a = string_table[j][number_column_nominal_thickness + 1:]
                        #     # минимальное значение
                        #     minim = float(100000)
                        #     # находим минимальное значение
                        #     for ii in a:
                        #         # если нашли число
                        #         if re.findall(r'\d', ii):
                        #             # меняем недопустимый символ ',' на '.'
                        #             if re.findall(r',', ii):
                        #                 ii = re.sub(',', r'.', ii)
                        #             a = float(ii)
                        #             if float(a) < float(minim):
                        #                 minim = a
                        #     index_minimum_thickness = 0
                        #     for ii in string_table[j]:
                        #         if str(minim) != ii:
                        #             index_minimum_thickness += 1
                        #         else:
                        #             break
                        #     # добавляем в список индекс минимального значения каждой строки
                        #     list_index_minimum_thickness.append(index_minimum_thickness)

                        # закрываем соединение
                        cur.close()

                        # переопределяем цвет для окрашивания столбца с номинальной толщиной
                        class ColorNominalThickness(QItemDelegate):
                            def __init__(self):
                                super().__init__()
                                self.filter = ''

                            def paint(self, painter, option, index):
                                # выбираем зелёный цвет для столбца с номинальной толщиной
                                painter.fillRect(option.rect, QColor(35, 198, 23, 180))
                                return QItemDelegate.paint(self, painter, option, index)

                        # создаём модель
                        color_nominal_thickness = ColorNominalThickness()
                        # окрашиваем столбец с номинальной толщиной в зелёный цвет
                        table_view[i].setItemDelegateForColumn(number_column_nominal_thickness, color_nominal_thickness)
                    scroll_area.show()

                # сообщение о том, что ничего не найдено
                else:
                    QMessageBox.information(
                        window,
                        'Внимание',
                        'Ничего не найдено!'
                    )

    # сообщение об ошибке, если в поле для поиска ничего не введено
    else:
        QMessageBox.information(
            window,
            'Внимание',
            'Вы не ввели ни номер линии, ни номер чертежа, ни номер репорта, ни номер word order для поиска данных'
        )


# функция отображения и повторного скрытия таблиц в frame
# l_t_v = list_table_view = список всех таблиц
# l_b_t = list_button_for_table = список всех номеров репортов
# y_1 - координата строки с номером репорта
# y_2 = y_1 + 20 - координата таблицы (20 - высота строки с номером репорта)
# l_h_t_v = list_height_table_view = список высот таблиц (строка с названием колонок + все строки таблицы)
def visible_table_view(l_t_v, l_b_t, l_h_t_v):
    # y_1 - координата первой строки номера репорта
    y_1 = 0
    # y_2 - координата первой таблицы
    y_2 = 20
    ii = 0
    # список новых координат номеров репортов
    position_y1 = []
    # список новых координат таблиц
    position_y2 = []
    # список нажатых кнопок
    list_button_for_table_true = []
    # список отжатых кнопок
    list_button_for_table_false = []
    for i in l_b_t:
        # если нажата
        if i.isChecked():
            list_button_for_table_true.append(ii)
        # если не нажата
        if not i.isChecked():
            list_button_for_table_false.append(ii)
        ii += 1
    # вычисляем новые координаты номеров репортов и таблиц в зависимости от списка нажатых (list_button_for_table_true)
    # и не нажатых (list_button_for_table_false) кнопок
    for i in range(len(l_h_t_v)):
        # если нажата кнопка номера репорта
        if list_button_for_table_true:
            # перебираем номера нажатых кнопок
            for ii in list_button_for_table_true:
                if i == ii:
                    # добавляем в список координату кнопки номера репорта
                    position_y1.append(y_1)
                    # меняем координату кнопки номера репорта, потому что она нажата и появляется таблица с данными
                    y_1 += 40 + l_h_t_v[i]
                    # добавляем в список координату таблицы с данными
                    position_y2.append(y_2)
                    # меняем координату кнопки номера репорта, потому что она нажата и появляется таблица с данными
                    y_2 += 40 + l_h_t_v[ii]
        # если НЕ нажата кнопка номера репорта
        if list_button_for_table_false:
            # перебираем номера НЕ нажатых кнопок
            for ii in list_button_for_table_false:
                if i == ii:
                    # добавляем в список Не нажатой координату кнопки номера репорта
                    position_y1.append(y_1)
                    y_1 += 20
                    # добавляем в список координату таблицы с данными при НЕ нажатой кнопки номера репорта
                    position_y2.append(y_2)
                    y_2 += 20
    # делаем таблицы видимыми или скрываем их в зависимости от статуса
    for b in list_button_for_table_true:
        l_b_t[b].move(0, position_y1[b])
        l_t_v[b].setGeometry((QRect(0, position_y2[b], 1460, l_h_t_v[b])))
        # делаем таблицу из списка видимой
        l_t_v[b].setVisible(True)
    for bb in list_button_for_table_false:
        l_b_t[bb].move(0, position_y1[bb])
        # делаем таблицу из списка снова скрытой
        l_t_v[bb].hide()


# нажатие на кнопку "Удалить"
def delete_report():
    pass


def log_out():
    # делаем НЕ активными кнопки "Добавить", "Редактировать", "Готово", "Удалить", "Выйти"
    button_repair.setDisabled(True)
    button_delete.setDisabled(True)
    button_log_out.setDisabled(True)
    # разблокируем кнопку "Войти"
    button_log_in.setDisabled(False)


# нажатие кнопки "Войти"
button_log_in.clicked.connect(log_in)
# нажатие на кнопку Enter когда фокус (каретка - мигающий символ "|") находится в поле для ввода логина
line_login.returnPressed.connect(log_in)
# нажатие на кнопку Enter когда фокус (каретка - мигающий символ "|") находится в поле для ввода пароля
line_password.returnPressed.connect(log_in)

# нажатие на кнопку "Добавить"
button_add.clicked.connect(add_tables)

# нажатие на кнопку "Поиск"
button_search.clicked.connect(search)
# нажатие на кнопку Enter когда фокус (каретка - мигающий символ "|") находится в поле для ввода номера линии, чертежа,
# номера репорта или work order
line_search.returnPressed.connect(search)

# нажатие на кнопку "Удалить"
button_delete.clicked.connect(delete_report)

# нажатие на кнопку "Выйти"
button_log_out.clicked.connect(log_out)


def main():
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
