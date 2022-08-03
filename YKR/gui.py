from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtSql import QSqlDatabase
from PyQt5.QtSql import QSqlQuery
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
window.resize(1524, 872)

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
line_search.setText('A1-3301-GA-006-1-A17-HC')

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
# делаем неактивной кнопку "Готово" до авторизации
button_ok.setDisabled(True)

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

# задаём поле для вывода данных из базы данных
tableView = QTableView(window)
tableView.setObjectName(u"tableView")
tableView.setGeometry(QRect(20, 140, 1481, 651))

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


# нажатие кнопки "Войти"
def log_in():
    if line_login.text() == 'admin' and line_password.text() == 'admin':
        # делаем активными кнопки "Добавить", "Редактировать", "Готово", "Удалить", "Выйти"
        button_add.setDisabled(False)
        button_repair.setDisabled(False)
        button_ok.setDisabled(False)
        button_delete.setDisabled(False)
        button_log_out.setDisabled(False)
        # очищаем поле ввода логина и пароля
        line_login.clear()
        line_password.clear()
        # блокируем кнопку "Войти"
        button_log_in.setDisabled(True)
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
            line_for_search = re.sub(r'"|\'\'', '', line_search.text())
        line_for_search = line_search.text()

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
            table_for_search = []
            for i in con.tables():
                # подключаемся в базе данных
                conn = sqlite3.connect('reports_db.sqlite')
                cur = conn.cursor()
                # перебираем список названий столбцов в таблице
                for k in cur.execute('SELECT * FROM {}'.format(i)).description:
                    # если 'Line' есть в названии столбца
                    if 'Line' in k:
                        # и если искомая линия есть в таблице, то добавляем имя таблицы в список
                        if cur.execute('SELECT Line FROM {} WHERE Line="{}"'.format(i, line_for_search)).fetchall():
                            table_for_search.append(i)
            # выводим данные в форму из найденных таблиц по номеру линии
            for i in table_for_search:
                # создаём модель
                sqm = QSqlQueryModel(parent=window)
                # создаём запрос
                sqm.setQuery('SELECT * FROM {} WHERE Line="{}"'.format(i, line_for_search),
                             db=QSqlDatabase('reports_db.sqlite'))
                # устанавливаем ширину столбцов под содержимое
                tableView.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                tableView.setModel(sqm)
        con.close()
    # сообщение об ошибке, если в поле для поиска ничего не введено
    else:
        QMessageBox.information(
            window,
            'Внимание',
            'Вы не ввели номер линии или чертежа для поиска данных'
        )


# нажатие кнопки "Войти"
button_log_in.clicked.connect(log_in)

# нажатие на кнопку "Добавить"
button_add.clicked.connect(add_tables)

# нажатие на кнопку "Поиск"
button_search.clicked.connect(search)


def log_out():
    # делаем НЕ активными кнопки "Добавить", "Редактировать", "Готово", "Удалить", "Выйти"
    button_add.setDisabled(True)
    button_repair.setDisabled(True)
    button_ok.setDisabled(True)
    button_delete.setDisabled(True)
    button_log_out.setDisabled(True)
    # разблокируем кнопку "Войти"
    button_log_in.setDisabled(False)


# нажатие кнопки "Выйти"
button_log_out.clicked.connect(log_out)

window.show()
sys.exit(app.exec_())
