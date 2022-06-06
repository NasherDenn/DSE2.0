from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication
import sys
import GUI_Front_End


class ExampleApp(QtWidgets.QMainWindow, GUI_Front_End.Ui_Dialog):
    def __init__(self, parent=None):
        super(ExampleApp, self).__init__(parent)
        self.setupUi(self)


def main():
    app = QApplication(sys.argv)
    form = ExampleApp()
    form.show()
    app.exec_()


if __name__ == '__main__':
    main()
