import sys

from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMessageBox, QMainWindow, QApplication, QFileDialog

from stock import Stock

form_class = uic.loadUiType("main_window.ui")[0]


class MyWindow(QMainWindow, form_class):
    DATE_FORMAT = 'yyyy-MM-ddThh:mm:ss.zzz'
    filename = ''

    def __init__(self):
        super().__init__()
        self.setupUi(self)

    @pyqtSlot()
    def file_button_clicked(self):
        self.filename = QFileDialog.getOpenFileName(self)[0]
        self.stock = Stock(self.filename)

    @pyqtSlot()
    def start_button_clicked(self):
        try:
            self.stock.get_roe()
            self.stock.get_dart()
            self.stock.write_xlsx()
        except Exception as e:
            QMessageBox.about(self, 'ERROR', 'FAIL!')
            return

        QMessageBox.about(self, 'INFO', 'SUCCESS!')

    @pyqtSlot()
    def update_button_clicked(self):
        QMessageBox.about(self, 'INFO', 'update is success!')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()