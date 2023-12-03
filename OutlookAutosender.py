from PyQt5 import QtWidgets
import os
import sys
from main_functions import *
from main_window import MainWindow

# pyinstaller --noconfirm --onedir --console --icon "C:/Users/dimas/OtpravkaRosreestr/icons8-carrier-pigeon-64.ico" --add-data "C:/Users/dimas/OtpravkaRosreestr/UI;UI" "C:/Users/dimas/OtpravkaRosreestr/OutlookAutosender.py"
# pyinstaller --noconfirm --onedir --console --windowed --icon "C:/Users/dimas/OtpravkaRosreestr/icons8-carrier-pigeon-64.ico" --add-data "C:/Users/dimas/OtpravkaRosreestr/UI;UI" --add-data "C:/Users/dimas/OtpravkaRosreestr/readme.txt;." "C:/Users/dimas/OtpravkaRosreestr/OutlookAutosender.py"


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_ui = MainWindow(config)
    sys.exit(app.exec_())

