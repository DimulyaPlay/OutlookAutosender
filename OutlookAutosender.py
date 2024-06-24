import logging
from PyQt5 import QtWidgets, QtGui
import os
import sys
from main_functions import *
from main_window import MainWindow
import traceback

# C:/Users/CourtUser/Desktop/release/OutlookAutosender/venv/Scripts/pyinstaller.exe --noconfirm --onedir --console --icon "C:/Users/CourtUser/Desktop/release/OutlookAutosender/icons8-carrier-pigeon-64.ico" --add-data "C:/Users/CourtUser/Desktop/release/OutlookAutosender/ReportsPrinted;ReportsPrinted" --add-data "C:/Users/CourtUser/Desktop/release/OutlookAutosender/wget;wget" --add-data "C:/Users/CourtUser/Desktop/release/OutlookAutosender/UI;UI" --add-data "C:/Users/CourtUser/Desktop/release/OutlookAutosender/readme.txt;." "C:/Users/CourtUser/Desktop/release/OutlookAutosender/OutlookAutosender.py"

logging.getLogger("PyQt5").setLevel(logging.WARNING)
log_path = os.path.join(os.path.dirname(sys.argv[0]), 'log.log')
logging.basicConfig(filename=log_path, level=logging.ERROR)


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
        main_ui = MainWindow(config)
        sys.exit(app.exec_())
    except:
        logging.exception('Непредвиденная ошибка во время работы')
