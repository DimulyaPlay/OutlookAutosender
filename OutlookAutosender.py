import logging
from PyQt5 import QtWidgets
import os
import sys
from main_functions import *
from main_window import MainWindow

# pyinstaller --noconfirm --onedir --console --icon "C:/Users/dimas/OtpravkaRosreestr/icons8-carrier-pigeon-64.ico" --add-data "C:/Users/dimas/OtpravkaRosreestr/UI;UI" "C:/Users/dimas/OtpravkaRosreestr/OutlookAutosender.py"
# C:/Users/CourtUser/Desktop/release/OutlookAutosender/venv/Scripts/pyinstaller.exe --noconfirm --onedir --console --windowed --icon "C:/Users/CourtUser/Desktop/release/OutlookAutosender/icons8-carrier-pigeon-64.ico" --add-data "C:/Users/CourtUser/Desktop/release/OutlookAutosender/UI;UI" --add-data "C:/Users/CourtUser/Desktop/release/OutlookAutosender/readme.txt;." "C:/Users/CourtUser/Desktop/release/OutlookAutosender/OutlookAutosender.py"

logging.getLogger("PyQt5").setLevel(logging.WARNING)
log_path = os.path.join(os.path.dirname(sys.argv[0]), 'log.log')
logging.basicConfig(filename=log_path, level=logging.ERROR)

if __name__ == '__main__':
    try:
        app = QtWidgets.QApplication(sys.argv)
        main_ui = MainWindow(config)
        sys.exit(app.exec_())
    except:
        logging.exception('Непредвиденная ошибка во время работы')
