import logging
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo
import os
import sys
from main_functions import *
from main_window import MainWindow
import traceback

# venv/Scripts/pyinstaller.exe --noconfirm --onedir --console --icon "icons8-carrier-pigeon-64.ico" --add-data "UI;UI" --add-data "readme.txt;." --add-data "update.cfg;." --add-data "update.exe;." OutlookAutosender.py
logging.getLogger("PyQt5").setLevel(logging.WARNING)
log_path = os.path.join(os.path.dirname(sys.argv[0]), 'log.log')
logging.basicConfig(filename=log_path, level=logging.ERROR)


if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

def exception_hook(exc_type, exc_value, exc_traceback):
    """
    Функция для перехвата исключений и отображения диалогового окна с ошибкой.
    """
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    error_dialog = QtWidgets.QErrorMessage()
    error_dialog.showMessage(error_msg)
    error_dialog.exec_()
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = exception_hook


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
