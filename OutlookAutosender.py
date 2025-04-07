from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo
import os
import sys
from main_functions import *
from main_window import MainWindow
import traceback

# venv/Scripts/pyinstaller.exe --noconfirm --onedir --console --icon "icons8-carrier-pigeon-64.ico" --add-data "UI;UI" --add-data "readme.txt;." --add-data "update.cfg;." --add-data "update.exe;." OutlookAutosender.py


if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

if getattr(sys, 'frozen', False):
    sys.stdout = open('console_output.log', 'a', buffering=1)
    sys.stderr = open('console_errors.log', 'a', buffering=1)


if __name__ == '__main__':
    try:
        app = QtWidgets.QApplication(sys.argv)
        translator = QTranslator()
        locale = QLocale.system().name()  # Получение системной локали
        path = QLibraryInfo.location(QLibraryInfo.TranslationsPath)  # Путь к переводам Qt
        translator.load("qtbase_" + locale, path)
        app.installTranslator(translator)
        main_ui = MainWindow(config)
        sys.exit(app.exec_())
    except:
        traceback.print_exc()
        logging.exception('Непредвиденная ошибка во время работы')
